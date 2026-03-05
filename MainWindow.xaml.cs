using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Media;
using System.Net;
using System.Net.Http;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace GoldDesktop;

public partial class MainWindow : Window
{
    private readonly HttpClient _httpClient;
    private readonly DispatcherTimer _timer;
    private readonly DispatcherTimer _predictTimer; // 每 2 分钟执行一次均线/上下线预测
    private readonly List<PricePoint> _history = new();
    private readonly PriceRepository _repository;
    private readonly DateTime _startTimeUtc = DateTime.UtcNow;
    private readonly TimeSpan _fetchInterval = TimeSpan.FromSeconds(15);
    // 内存中仅保留最近 15 分钟的数据点
    private readonly TimeSpan _historyWindow = TimeSpan.FromMinutes(15);
    private readonly TimeSpan _predictHorizon = TimeSpan.FromMinutes(30);
    private readonly TimeSpan _predictInterval = TimeSpan.FromMinutes(2);

    /// <summary> true = 人民币/克，false = 美元/盎司。图表与预测均按此单位显示。 </summary>
    private bool _useCny;
    /// <summary> true 表示锁定横坐标，刷新时不再自动重置 X 轴范围。 </summary>
    private bool _lockXAxis;
    /// <summary> Monte Carlo 随机数生成器。 </summary>
    private static readonly Random _rng = new();
    /// <summary> 图表用的最近一次 Monte Carlo 区间带（按 1 分钟步长），仅在每分钟预测时更新。 </summary>
    private List<(DateTime TimeLocal, double Upper, double Lower)>? _latestMcBandForChart;
    /// <summary> 图表用的最近一次 Monte Carlo 代表路径（蓝/绿/红三条），仅在每分钟预测时更新。 </summary>
    private List<List<(DateTime TimeLocal, double Price)>>? _latestMcSamplePathsForChart;

    /// <summary> 价格提醒：已触发过“低于”提醒，关闭弹窗后等价格回到阈值以上再重置。 </summary>
    private bool _lastAlertedBelow;
    /// <summary> 价格提醒：已触发过“高于”提醒，关闭弹窗后等价格回到阈值以下再重置。 </summary>
    private bool _lastAlertedAbove;

    /// <summary> 价格提醒时播放的本地音效文件名。 </summary>
    private const string AlertSoundFileName = "1.wav";
    private MediaPlayer? _alertPlayer;
    /// <summary> 为 true 时提示音循环播放，关闭弹窗后设为 false 并停止。 </summary>
    private bool _alertSoundLooping;

    public MainWindow()
    {
        InitializeComponent();

        var dbPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "gold.db");
        _repository = new PriceRepository(dbPath);
        _repository.ClearAll();

        var handler = new HttpClientHandler
        {
            AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate
        };
        _httpClient = new HttpClient(handler);

        _timer = new DispatcherTimer { Interval = _fetchInterval };
        _timer.Tick += async (_, _) => await FetchAndUpdateAsync();
        _timer.Start();

        _predictTimer = new DispatcherTimer { Interval = _predictInterval };
        _predictTimer.Tick += (_, _) => RunPredictionAndUpdateBands();
        _predictTimer.Start();

        Loaded += async (_, _) => { await FetchAndUpdateAsync(); };
    }

    private async System.Threading.Tasks.Task FetchAndUpdateAsync()
    {
        TxtTime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
        try
        {
            TxtStatus.Text = "正在获取实时价格...";
            var result = await FetchGoldPriceAsync();
            if (result is null)
            {
                TxtStatus.Text = "获取价格失败。";
                return;
            }

            if (result.UsdDetails is not null)
            {
                var u = result.UsdDetails;
                TxtPriceUsd.Text = u.CurrentPrice.ToString("F2");
                TxtChangeUsd.Text =
                    $"{u.ChangeAmount:+0.00;-0.00;0.00}  ({u.ChangePercent:+0.00;-0.00;0.00}%)";
            }
            else
            {
                TxtPriceUsd.Text = "--";
                TxtChangeUsd.Text = "";
            }

            if (result.CnyDetails is not null)
            {
                var c = result.CnyDetails;
                TxtPriceCny.Text = c.CurrentPrice.ToString("F2");
                TxtChangeCny.Text =
                    $"{c.ChangeAmount:+0.00;-0.00;0.00}  ({c.ChangePercent:+0.00;-0.00;0.00}%)";
            }
            else
            {
                TxtPriceCny.Text = "--";
                TxtChangeCny.Text = "";
            }

            var usd = result.UsdPrice ?? 0;
            var cny = result.CnyPrice ?? 0;
            if (usd <= 0 && cny <= 0)
            {
                TxtStatus.Text = "未获取到有效价格。";
                return;
            }

            var now = DateTime.UtcNow;
            _history.Add(new PricePoint(now, usd, cny));
            _history.RemoveAll(p => now - p.Time > _historyWindow);

            // 保存到本地 SQLite
            _repository.SaveTick(now, result.UsdPrice, result.CnyPrice);

            UpdateKline();

            // 价格提醒：与当前显示单位一致，低于/高于设定值时播放提示音并弹窗，关闭弹窗后消除本次提醒
            var currentPriceForAlert = _useCny ? cny : usd;
            if (currentPriceForAlert > 0)
                CheckPriceAlert(currentPriceForAlert);

            TxtStatus.Text = result.Summary;
        }
        catch (Exception ex)
        {
            TxtStatus.Text = $"发生异常：{ex.Message}";
        }
    }

    private async System.Threading.Tasks.Task<GoldPriceResult?> FetchGoldPriceAsync()
    {
        var headers = new Dictionary<string, string>
        {
            { "authority", "api.jijinhao.com" },
            { "accept", "*/*" },
            { "accept-language", "zh-CN,zh;q=0.9" },
            { "referer", "https://quote.cngold.org/" },
            { "sec-ch-ua", "\"Not)A;Brand\";v=\"24\", \"Chromium\";v=\"116\"" },
            { "sec-ch-ua-mobile", "?0" },
            { "sec-ch-ua-platform", "\"Windows\"" },
            { "sec-fetch-dest", "script" },
            { "sec-fetch-mode", "no-cors" },
            { "sec-fetch-site", "cross-site" },
            {
                "user-agent",
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/116.0.5845.97 Safari/537.36 Core/1.116.541.400 QQBrowser/19.4.6579.400"
            }
        };

        // 时间戳参数，避免缓存，例如：&_={1772546532698}
        var ts = DateTimeOffset.UtcNow.ToUnixTimeMilliseconds();
        var urlUsd =
            $"https://api.jijinhao.com/sQuoteCenter/realTime.htm?code=JO_92233&_={ts}";
        var urlCny =
            $"https://api.jijinhao.com/sQuoteCenter/realTime.htm?code=JO_92233&isCalc=true&_={ts}";

        var reqUsd = new HttpRequestMessage
        {
            Method = HttpMethod.Get,
            RequestUri = new Uri(urlUsd)
        };
        foreach (var kv in headers) reqUsd.Headers.TryAddWithoutValidation(kv.Key, kv.Value);

        var reqCny = new HttpRequestMessage
        {
            Method = HttpMethod.Get,
            RequestUri = new Uri(urlCny)
        };
        foreach (var kv in headers) reqCny.Headers.TryAddWithoutValidation(kv.Key, kv.Value);

        using var respUsd = await _httpClient.SendAsync(reqUsd);
        respUsd.EnsureSuccessStatusCode();
        var dataUsd = await respUsd.Content.ReadAsStringAsync();

        using var respCny = await _httpClient.SendAsync(reqCny);
        respCny.EnsureSuccessStatusCode();
        var dataCny = await respCny.Content.ReadAsStringAsync();

        GoldPriceDetail? ParseOne(string raw, double openPrice, string unit)
        {
            const string marker = "var hq_str = ";
            var idx = raw.IndexOf(marker, StringComparison.Ordinal);
            if (idx < 0) return null;

            var part = raw[(idx + marker.Length)..].Trim();
            if (!part.StartsWith("\"", StringComparison.Ordinal)) return null;
            var secondQuoteIdx = part.IndexOf('"', 1);
            if (secondQuoteIdx <= 1) return null;

            var hqStr = part[1..secondQuoteIdx];
            var parts = hqStr.Split(',');
            if (parts.Length < 9) return null;

            if (!double.TryParse(parts[2], out var yesterday) ||
                !double.TryParse(parts[3], out var current) ||
                !double.TryParse(parts[4], out var high) ||
                !double.TryParse(parts[5], out var low))
                return null;

            var change = current - yesterday;
            var changePercent = Math.Abs(yesterday) > 1e-9 ? change / yesterday * 100 : 0;

            return new GoldPriceDetail
            {
                Name = parts[0],
                YesterdayClose = yesterday,
                CurrentPrice = current,
                HighPrice = high,
                LowPrice = low,
                OpenPrice = openPrice,
                Unit = unit,
                ChangeAmount = change,
                ChangePercent = changePercent
            };
        }

        var usdDetail = ParseOne(dataUsd, 4994.17, "美元/盎司");
        var cnyDetail = ParseOne(dataCny, 1109.6888, "元/克");

        if (usdDetail is null && cnyDetail is null) return null;

        return new GoldPriceResult
        {
            UsdDetails = usdDetail,
            CnyDetails = cnyDetail,
            UsdPrice = usdDetail?.CurrentPrice,
            CnyPrice = cnyDetail?.CurrentPrice
        };
    }

    private void UpdateKline()
    {
        // 使用最近一段时间（与 _historyWindow 一致）的数据作为折线图数据源
        var cutoff = _startTimeUtc > DateTime.UtcNow - _historyWindow
            ? _startTimeUtc
            : DateTime.UtcNow - _historyWindow;
        var points = _history.Where(p => p.Time >= cutoff)
            .OrderBy(p => p.Time)
            .ToList();
        if (points.Count < 2 || KlinePlot == null)
        {
            if (KlinePlot != null) KlinePlot.Model = null;
            return;
        }

        // 先转换为本地时间 + 当前单位价格，便于后续统一使用
        var localPoints = points
            .Select(p => new
            {
                TimeLocal = p.Time.ToLocalTime(),
                Price = p.Price(_useCny)
            })
            .Where(x => x.Price > 0)
            .ToList();
        if (localPoints.Count < 2)
        {
            KlinePlot.Model = null;
            return;
        }

        // 如果已锁定横坐标：只刷新数据点，不改变当前 X/Y 轴范围，保持用户视角不动
        if (_lockXAxis && KlinePlot.Model is { } existingModel)
        {
            var lineExisting = existingModel.Series.OfType<OxyPlot.Series.LineSeries>().FirstOrDefault();
            var labelExisting = existingModel.Series.OfType<OxyPlot.Series.ScatterSeries>().FirstOrDefault();
            if (lineExisting != null && labelExisting != null)
            {
                lineExisting.Points.Clear();
                labelExisting.Points.Clear();
                foreach (var p in localPoints)
                {
                    var x = OxyPlot.Axes.DateTimeAxis.ToDouble(p.TimeLocal);
                    lineExisting.Points.Add(new OxyPlot.DataPoint(x, p.Price));
                    labelExisting.Points.Add(new OxyPlot.Series.ScatterPoint(x, p.Price));
                }
                KlinePlot.InvalidatePlot(true);
                return;
            }
        }

        // 否则处于自动跟随模式：每次有新点位时，以最新点为右端，显示最近约一小时窗口
        var unitLabel = _useCny ? "元/克" : "美元/盎司";
        var model = new OxyPlot.PlotModel { Title = $"黄金价格折线图（最近 1 小时，{unitLabel}）" };

        var minPrice = localPoints.Min(p => p.Price);
        var maxPrice = localPoints.Max(p => p.Price);
        var range = Math.Max(maxPrice - minPrice, 0.01);
        var margin = range * 0.2;

        var lastTimeLocal = localPoints.Last().TimeLocal;
        // 横坐标：仅显示最近 15 分钟中的“较新的”一段，使新增点每次看起来向左多移动约 2 个 15s 单位
        // 基本思路：总共只保留 15 分钟数据，但可视窗口略窄一点（比如 14.5 分钟），让旧数据更快从左边滑出
        var step = TimeSpan.FromSeconds(15);              // 采样间隔
        var visibleWindow = _historyWindow - step * 2;    // 比 15 分钟略少 30 秒
        if (visibleWindow <= TimeSpan.Zero) visibleWindow = _historyWindow;
        var xMin = OxyPlot.Axes.DateTimeAxis.ToDouble(lastTimeLocal - visibleWindow);
        var xMax = OxyPlot.Axes.DateTimeAxis.ToDouble(lastTimeLocal + TimeSpan.FromMinutes(7));

        var xAxis = new OxyPlot.Axes.DateTimeAxis
        {
            Position = OxyPlot.Axes.AxisPosition.Bottom,
            StringFormat = "HH:mm:ss",
            IntervalType = OxyPlot.Axes.DateTimeIntervalType.Seconds,
            MajorStep = TimeSpan.FromSeconds(15).TotalDays,
            Minimum = xMin,
            Maximum = xMax,
            MinimumPadding = 0,
            MaximumPadding = 0,
            IsZoomEnabled = true,
            IsPanEnabled = true
        };
        var yAxis = new OxyPlot.Axes.LinearAxis
        {
            Position = OxyPlot.Axes.AxisPosition.Left,
            Minimum = minPrice - margin,
            Maximum = maxPrice + margin,
            IsZoomEnabled = true,
            IsPanEnabled = true
        };
        model.Axes.Add(xAxis);
        model.Axes.Add(yAxis);

        var lineSeries = new OxyPlot.Series.LineSeries
        {
            Color = OxyPlot.OxyColors.Goldenrod,
            StrokeThickness = 2,
            MarkerType = OxyPlot.MarkerType.Circle,
            MarkerSize = 3,
            MarkerFill = OxyPlot.OxyColors.White,
            MarkerStroke = OxyPlot.OxyColors.Goldenrod
        };

        var labelSeries = new OxyPlot.Series.ScatterSeries
        {
            MarkerType = OxyPlot.MarkerType.Circle,
            MarkerSize = 0,
            LabelFormatString = "{1:0.00}",
            TextColor = OxyPlot.OxyColors.DimGray
        };

        foreach (var p in localPoints)
        {
            var x = OxyPlot.Axes.DateTimeAxis.ToDouble(p.TimeLocal);
            lineSeries.Points.Add(new OxyPlot.DataPoint(x, p.Price));
            labelSeries.Points.Add(new OxyPlot.Series.ScatterPoint(x, p.Price));
        }

        // 先叠加 Monte Carlo 未来 30 分钟价格区间阴影带（5%~95% 区间），再绘制折线。
        // 注意：区间带仅在每分钟的预测计时器中刷新，这里只读取缓存，避免每 15 秒重新模拟，增强参考意义。
        var mcBand = _latestMcBandForChart;
        if (mcBand is { Count: > 0 })
        {
            var area = new OxyPlot.Series.AreaSeries
            {
                Color = OxyPlot.OxyColors.Transparent,
                Fill = OxyPlot.OxyColor.FromAColor(40, OxyPlot.OxyColors.SteelBlue),
                StrokeThickness = 1,
                LineStyle = OxyPlot.LineStyle.Dash,
                Title = "未来区间 (MC)"
            };
            foreach (var p in mcBand)
            {
                var x = OxyPlot.Axes.DateTimeAxis.ToDouble(p.TimeLocal);
                area.Points.Add(new OxyPlot.DataPoint(x, p.Upper));
                area.Points2.Add(new OxyPlot.DataPoint(x, p.Lower));
            }
            model.Series.Add(area);
        }

        // 三条代表性 Monte Carlo 未来路径：蓝/绿/红三条虚线，表示偏低/中性/偏高三种可能走势（同样使用最近一次缓存）
        var samplePaths = _latestMcSamplePathsForChart ?? new List<List<(DateTime TimeLocal, double Price)>>();
        var sampleColors = new[]
        {
            OxyPlot.OxyColors.SteelBlue,
            OxyPlot.OxyColors.ForestGreen,
            OxyPlot.OxyColors.Red
        };
        var sampleTitles = new[]
        {
            "MC 走势(偏低)",
            "MC 走势(中性)",
            "MC 走势(偏高)"
        };
        for (var i = 0; i < Math.Min(3, samplePaths.Count); i++)
        {
            var pathSeries = new OxyPlot.Series.LineSeries
            {
                Color = sampleColors[i],
                LineStyle = OxyPlot.LineStyle.Dot,
                StrokeThickness = 1.5,
                Title = sampleTitles[i]
            };
            foreach (var p in samplePaths[i])
            {
                var x = OxyPlot.Axes.DateTimeAxis.ToDouble(p.TimeLocal);
                pathSeries.Points.Add(new OxyPlot.DataPoint(x, p.Price));
            }
            model.Series.Add(pathSeries);
        }

        model.Series.Add(lineSeries);
        model.Series.Add(labelSeries);

        KlinePlot.Model = model;
    }

    /// <summary>
    /// 每 30 秒执行：基于历史数据计算未来 30 分钟的均线、上线、下线并更新界面。
    /// </summary>
    private void RunPredictionAndUpdateBands()
    {
        // 限制列表长度，避免无限增长（保留最近约 10 组预测）
        const int maxItems = 10 * 10; // 每组大约 8~10 行
        if (LstForecast.Items.Count > maxItems)
        {
            LstForecast.Items.Clear();
        }

        if (_history.Count < 10)
        {
            LstForecast.Items.Add($"[{DateTime.Now:HH:mm:ss}] 数据收集中（当前 {_history.Count} 个点），稍后给出均线/上下线预测。");
            LstForecast.Items.Add("※ 仅供参考，不构成投资建议。");
            LstForecast.Items.Add("");
            return;
        }

        var (points, guidance) = PredictWithBands(_history, _predictHorizon, _useCny);
        if (points.Count == 0)
        {
            LstForecast.Items.Add($"[{DateTime.Now:HH:mm:ss}] 暂无预测数据。");
            LstForecast.Items.Add("");
            return;
        }

        var lastPrice = _history.Last().Price(_useCny);
        var unitLabel = _useCny ? "元/克" : "美元/盎司";
        LstForecast.Items.Add($"========== {DateTime.Now:HH:mm:ss} 新一轮预测 ({unitLabel}) ==========");
        LstForecast.Items.Add($"【当前价】{lastPrice:F2} {unitLabel}");
        foreach (var p in points)
        {
            LstForecast.Items.Add(
                $"  未来 {p.Offset.TotalMinutes,2:F0} 分钟  均线: {p.Ma:F2}  上线: {p.Upper:F2}  下线: {p.Lower:F2}");
        }
        LstForecast.Items.Add($"【指导意见】{guidance}");

        // Monte Carlo 模拟未来区间（更接近量化回测工具的做法）
        var mcPoints = PredictWithMonteCarlo(_history, _predictHorizon, _useCny, 300);
        if (mcPoints.Count > 0)
        {
            LstForecast.Items.Add("");
            LstForecast.Items.Add($"---- Monte Carlo 模拟（{mcPoints.Count}×4 节点，{unitLabel}） ----");
            foreach (var p in mcPoints)
            {
                LstForecast.Items.Add(
                    $"  未来 {p.Offset.TotalMinutes,2:F0} 分钟  中值: {p.Ma:F2}  95%上线: {p.Upper:F2}  5%下线: {p.Lower:F2}");
            }
        }

        // 更新图表用的 Monte Carlo 缓存（每 1 分钟一次）
        _latestMcBandForChart = PredictMonteCarloForChart(_history, _useCny, _predictHorizon, 200, TimeSpan.FromMinutes(1));
        _latestMcSamplePathsForChart = PredictMonteCarloSamplePathsForChart(_history, _useCny, _predictHorizon, 300, TimeSpan.FromMinutes(1));

        LstForecast.Items.Add("※ 仅供参考，不构成投资建议。");
        LstForecast.Items.Add("");
    }

    /// <summary>
    /// 三点一线方式：用最近 3 个历史点拟合趋势线作为均线，
    /// 用最近 30 分钟的波动率计算上线/下线趋势线，并生成简要指导意见。
    /// </summary>
    private static (List<ForecastBandPoint> Points, string Guidance) PredictWithBands(
        List<PricePoint> history,
        TimeSpan horizon,
        bool useCny)
    {
        if (history.Count < 3)
        {
            return (new List<ForecastBandPoint>(), "历史点不足 3 个，无法计算三点一线趋势。");
        }

        // 按所选单位取价格序列（过滤无效点）
        var working = history.Select(p => (p.Time, Price: p.Price(useCny))).Where(x => x.Price > 0).ToList();
        if (working.Count < 3)
        {
            return (new List<ForecastBandPoint>(), "当前单位有效历史点不足 3 个。");
        }

        var lastTime = working.Last().Time;
        var recentSpan = TimeSpan.FromMinutes(30);
        var cutoff = lastTime - recentSpan;
        var recent = working.Where(p => p.Time >= cutoff).ToList();
        if (recent.Count < 5) recent = working;

        var trendPoints = working.Skip(Math.Max(0, working.Count - 3)).ToList();
        var t0 = trendPoints.First().Time;
        var xsTrend = trendPoints.Select(p => (p.Time - t0).TotalSeconds).ToArray();
        var ysTrend = trendPoints.Select(p => p.Price).ToArray();

        LinearRegressionSimple(xsTrend, ysTrend, out var aTrend, out var bTrend);

        // 用最近一段历史的标准差估计带宽
        var meanRecent = recent.Average(p => p.Price);
        var varRecent = recent.Select(p => (p.Price - meanRecent) * (p.Price - meanRecent)).Average();
        var stdRecent = Math.Sqrt(varRecent);
        if (stdRecent < 0.01) stdRecent = 0.01;
        var bandHalfWidth = 2 * stdRecent; // 上下线相对均线的距离

        var checkpoints = new[]
        {
            TimeSpan.FromMinutes(5),
            TimeSpan.FromMinutes(10),
            TimeSpan.FromMinutes(15),
            TimeSpan.FromMinutes(30)
        }.Where(t => t <= horizon).ToArray();

        var points = new List<ForecastBandPoint>();
        foreach (var offset in checkpoints)
        {
            var futureTime = lastTime + offset;
            var x = (futureTime - t0).TotalSeconds;
            var ma = aTrend + bTrend * x; // 三点一线趋势均线
            var upper = ma + bandHalfWidth;
            var lower = ma - bandHalfWidth;
            points.Add(new ForecastBandPoint(offset, ma, upper, lower));
        }

        var lastPrice = working.Last().Price;
        var slope = bTrend; // 单位：价格/秒，可粗略视为趋势方向
        var maNow = aTrend + bTrend * (lastTime - t0).TotalSeconds;
        var bandWidth = points.Count > 0 ? (points[0].Upper - points[0].Lower) : 0;

        // 简要指导意见
        var guidance = BuildGuidance(lastPrice, maNow, slope, bandWidth, points);
        return (points, guidance);
    }

    /// <summary>
    /// Monte Carlo 方式：根据最近一段历史收益分布，模拟多条未来路径，取中位数/5%/95% 作为未来区间。
    /// </summary>
    private static List<ForecastBandPoint> PredictWithMonteCarlo(
        List<PricePoint> history,
        TimeSpan horizon,
        bool useCny,
        int pathCount)
    {
        // 至少需要若干个点来估计收益分布
        if (history.Count < 10) return new List<ForecastBandPoint>();

        var working = history
            .Select(p => (p.Time, Price: p.Price(useCny)))
            .Where(x => x.Price > 0)
            .OrderBy(x => x.Time)
            .ToList();
        if (working.Count < 10) return new List<ForecastBandPoint>();

        // 计算对数收益 r_t = ln(P_t / P_{t-1})
        var returns = new List<double>();
        for (var i = 1; i < working.Count; i++)
        {
            var pPrev = working[i - 1].Price;
            var pNow = working[i].Price;
            if (pPrev <= 0 || pNow <= 0) continue;
            returns.Add(Math.Log(pNow / pPrev));
        }
        if (returns.Count < 5) return new List<ForecastBandPoint>();

        var avgR = returns.Average();
        var last = working.Last();
        var lastPrice = last.Price;

        var checkpoints = new[]
        {
            TimeSpan.FromMinutes(5),
            TimeSpan.FromMinutes(10),
            TimeSpan.FromMinutes(15),
            TimeSpan.FromMinutes(30)
        }.Where(t => t <= horizon).ToArray();
        if (checkpoints.Length == 0) return new List<ForecastBandPoint>();

        // 假设采样间隔近似均匀，按历史平均间隔估算每步为 dtStep
        var avgDtSeconds = working.Zip(working.Skip(1), (a, b) => (b.Time - a.Time).TotalSeconds).Average();
        if (avgDtSeconds <= 0) avgDtSeconds = 15; // 兜底用 15 秒

        var result = new List<ForecastBandPoint>();

        foreach (var offset in checkpoints)
        {
            var steps = Math.Max(1, (int)Math.Round(offset.TotalSeconds / avgDtSeconds));
            var terminalPrices = new double[pathCount];

            for (var p = 0; p < pathCount; p++)
            {
                var price = lastPrice;
                for (var s = 0; s < steps; s++)
                {
                    // 简单方法：从历史收益中随机抽样一条（bootstrap），避免假设正态
                    var r = returns[_rng.Next(returns.Count)];
                    // 可加入平均漂移：r += avgR;
                    price *= Math.Exp(r);
                }
                terminalPrices[p] = price;
            }

            Array.Sort(terminalPrices);
            double Quantile(double q)
            {
                var idx = (int)Math.Round(q * (terminalPrices.Length - 1));
                idx = Math.Clamp(idx, 0, terminalPrices.Length - 1);
                return terminalPrices[idx];
            }

            var median = Quantile(0.5);
            var p5 = Quantile(0.05);
            var p95 = Quantile(0.95);

            result.Add(new ForecastBandPoint(offset, median, p95, p5));
        }

        return result;
    }

    /// <summary>
    /// Monte Carlo：为图表生成曲线用的区间带（每 step 一个时间点），返回本地时间 + 上下沿。
    /// </summary>
    private static List<(DateTime TimeLocal, double Upper, double Lower)> PredictMonteCarloForChart(
        List<PricePoint> history,
        bool useCny,
        TimeSpan horizon,
        int pathCount,
        TimeSpan step)
    {
        var result = new List<(DateTime, double, double)>();
        if (history.Count < 10 || step <= TimeSpan.Zero) return result;

        var working = history
            .Select(p => (p.Time, Price: p.Price(useCny)))
            .Where(x => x.Price > 0)
            .OrderBy(x => x.Time)
            .ToList();
        if (working.Count < 10) return result;

        var returns = new List<double>();
        for (var i = 1; i < working.Count; i++)
        {
            var pPrev = working[i - 1].Price;
            var pNow = working[i].Price;
            if (pPrev <= 0 || pNow <= 0) continue;
            returns.Add(Math.Log(pNow / pPrev));
        }
        if (returns.Count < 5) return result;

        var last = working.Last();
        var lastPrice = last.Price;
        var lastTime = last.Time;

        var steps = Math.Max(1, (int)Math.Round(horizon.TotalSeconds / step.TotalSeconds));
        var allPaths = new double[pathCount, steps];

        for (var p = 0; p < pathCount; p++)
        {
            var price = lastPrice;
            for (var s = 0; s < steps; s++)
            {
                var r = returns[_rng.Next(returns.Count)];
                price *= Math.Exp(r);
                allPaths[p, s] = price;
            }
        }

        for (var s = 0; s < steps; s++)
        {
            var slice = new double[pathCount];
            for (var p = 0; p < pathCount; p++)
                slice[p] = allPaths[p, s];
            Array.Sort(slice);

            double Quantile(double q)
            {
                var idx = (int)Math.Round(q * (slice.Length - 1));
                idx = Math.Clamp(idx, 0, slice.Length - 1);
                return slice[idx];
            }

            var p5 = Quantile(0.05);
            var p95 = Quantile(0.95);
            var tLocal = (lastTime + TimeSpan.FromTicks(step.Ticks * (s + 1))).ToLocalTime();
            result.Add((tLocal, p95, p5));
        }

        // 为了让阴影带与当前折线自然衔接，补一个锚点：当前时刻，上下沿都等于当前价格
        var anchorTimeLocal = lastTime.ToLocalTime();
        result.Insert(0, (anchorTimeLocal, lastPrice, lastPrice));

        return result;
    }

    /// <summary>
    /// Monte Carlo：挑选 3 条代表性未来路径（偏低/中性/偏高），用于图表上画蓝/绿/红三条猜测走势。
    /// </summary>
    private static List<List<(DateTime TimeLocal, double Price)>> PredictMonteCarloSamplePathsForChart(
        List<PricePoint> history,
        bool useCny,
        TimeSpan horizon,
        int pathCount,
        TimeSpan step)
    {
        var result = new List<List<(DateTime, double)>>();
        if (history.Count < 10 || step <= TimeSpan.Zero) return result;

        var working = history
            .Select(p => (p.Time, Price: p.Price(useCny)))
            .Where(x => x.Price > 0)
            .OrderBy(x => x.Time)
            .ToList();
        if (working.Count < 10) return result;

        var returns = new List<double>();
        for (var i = 1; i < working.Count; i++)
        {
            var pPrev = working[i - 1].Price;
            var pNow = working[i].Price;
            if (pPrev <= 0 || pNow <= 0) continue;
            returns.Add(Math.Log(pNow / pPrev));
        }
        if (returns.Count < 5) return result;

        var last = working.Last();
        var lastPrice = last.Price;
        var lastTime = last.Time;

        var steps = Math.Max(1, (int)Math.Round(horizon.TotalSeconds / step.TotalSeconds));
        var allPaths = new double[pathCount, steps];
        var terminals = new (double Price, int Index)[pathCount];

        for (var p = 0; p < pathCount; p++)
        {
            var price = lastPrice;
            for (var s = 0; s < steps; s++)
            {
                var r = returns[_rng.Next(returns.Count)];
                price *= Math.Exp(r);
                allPaths[p, s] = price;
            }
            terminals[p] = (allPaths[p, steps - 1], p);
        }

        // 按终点价格排序，选偏低(30%)、中性(50%)、偏高(70%) 三条路径
        Array.Sort(terminals, (a, b) => a.Price.CompareTo(b.Price));
        int ClampIndex(double q)
        {
            var idx = (int)Math.Round(q * (terminals.Length - 1));
            return Math.Clamp(idx, 0, terminals.Length - 1);
        }

        var idxLow = terminals[ClampIndex(0.3)].Index;
        var idxMid = terminals[ClampIndex(0.5)].Index;
        var idxHigh = terminals[ClampIndex(0.7)].Index;
        var indices = new[] { idxLow, idxMid, idxHigh };

        foreach (var pathIndex in indices)
        {
            var path = new List<(DateTime, double)>();
            for (var s = 0; s < steps; s++)
            {
                var tLocal = (lastTime + TimeSpan.FromTicks(step.Ticks * (s + 1))).ToLocalTime();
                path.Add((tLocal, allPaths[pathIndex, s]));
            }
            result.Add(path);
        }

        return result;
    }

    private static string BuildGuidance(
        double currentPrice,
        double maNow,
        double slopePerSec,
        double bandWidth,
        List<ForecastBandPoint> points)
    {
        var trend = slopePerSec > 0.0001 ? "偏多" : slopePerSec < -0.0001 ? "偏空" : "震荡";
        var pos = currentPrice > maNow ? "在均线上方" : currentPrice < maNow ? "在均线下方" : "在均线附近";
        var first = points.Count > 0 ? points[0] : null;
        var last = points.Count > 0 ? points[^1] : null;

        if (first is null || last is null)
            return $"趋势{trend}，当前价格{pos}。";

        var upperNear = Math.Abs(currentPrice - first.Upper) < bandWidth * 0.2;
        var lowerNear = Math.Abs(currentPrice - first.Lower) < bandWidth * 0.2;

        if (trend == "偏多" && upperNear)
            return "趋势偏多，价格接近上线，注意短线压力与回调风险，可关注均线支撑。";
        if (trend == "偏空" && lowerNear)
            return "趋势偏空，价格接近下线，注意短线支撑与反弹可能，可关注均线压力。";
        if (trend == "偏多")
            return "趋势偏多，价格在均线之上，上线为压力参考、下线为支撑参考，可沿均线顺势观察。";
        if (trend == "偏空")
            return "趋势偏空，价格在均线之下，下线为支撑参考、上线为压力参考，可沿均线顺势观察。";
        return "趋势震荡，价格在均线附近，建议在上下线区间内高抛低吸、设好止损。";
    }

    /// <summary>
    /// 简单线性回归 y = a + b*x（不返回方差等），用于三点一线趋势。
    /// </summary>
    private static void LinearRegressionSimple(double[] xs, double[] ys, out double a, out double b)
    {
        var n = xs.Length;
        if (n == 0)
        {
            a = 0;
            b = 0;
            return;
        }

        var sumX = xs.Sum();
        var sumY = ys.Sum();
        var sumXX = xs.Select(x => x * x).Sum();
        var sumXY = xs.Zip(ys, (x, y) => x * y).Sum();

        var denominator = n * sumXX - sumX * sumX;
        if (Math.Abs(denominator) < 1e-9)
        {
            a = sumY / n;
            b = 0;
            return;
        }

        b = (n * sumXY - sumX * sumY) / denominator;
        a = (sumY - b * sumX) / n;
    }

    /// <summary>
    /// 在程序目录及上级目录查找 1.wav，返回第一个存在的完整路径，否则返回 null。
    /// </summary>
    private static string? FindAlertSoundPath()
    {
        var baseDir = AppDomain.CurrentDomain.BaseDirectory;
        for (var dir = baseDir; !string.IsNullOrEmpty(dir); dir = System.IO.Path.GetDirectoryName(dir))
        {
            var path = System.IO.Path.Combine(dir, AlertSoundFileName);
            if (System.IO.File.Exists(path))
                return System.IO.Path.GetFullPath(path);
        }
        return null;
    }

    /// <summary>
    /// 开始循环播放价格提醒音效 1.wav，直到调用 StopAlertSound。文件不存在或失败则系统 Beep 一次。
    /// </summary>
    private void PlayAlertSoundLooping()
    {
        try
        {
            var path = FindAlertSoundPath();
            if (string.IsNullOrEmpty(path))
            {
                SystemSounds.Beep.Play();
                return;
            }
            if (_alertPlayer == null)
            {
                _alertPlayer = new MediaPlayer();
                _alertPlayer.MediaFailed += (_, _) => SystemSounds.Beep.Play();
                _alertPlayer.MediaEnded += (_, _) =>
                {
                    if (_alertSoundLooping && _alertPlayer != null)
                    {
                        _alertPlayer.Position = TimeSpan.Zero;
                        _alertPlayer.Play();
                    }
                };
            }
            _alertSoundLooping = true;
            // MediaPlayer 需要 file:// 形式的 URI，且路径用正斜杠
            var uri = new Uri("file:///" + path.Replace("\\", "/"));
            _alertPlayer.Open(uri);
            _alertPlayer.Position = TimeSpan.Zero;
            _alertPlayer.Volume = 1.0;
            _alertPlayer.Play();
        }
        catch
        {
            SystemSounds.Beep.Play();
        }
    }

    /// <summary>
    /// 停止提醒音循环（关闭弹窗后调用）。
    /// </summary>
    private void StopAlertSound()
    {
        _alertSoundLooping = false;
        _alertPlayer?.Stop();
    }

    /// <summary>
    /// 检查价格是否触及提醒阈值，与当前显示单位一致。触及则播放提示音并弹窗，关闭弹窗后消除提醒；价格回到区间内则重置以便再次提醒。
    /// </summary>
    private void CheckPriceAlert(double currentPrice)
    {
        if (TxtAlertBelow == null || TxtAlertAbove == null) return;

        double? below = double.TryParse(TxtAlertBelow.Text?.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var b) && b > 0 ? b : null;
        double? above = double.TryParse(TxtAlertAbove.Text?.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out var a) && a > 0 ? a : null;

        var unitLabel = _useCny ? "元/克" : "美元/盎司";

        if (below.HasValue && currentPrice > below.Value)
            _lastAlertedBelow = false;
        if (above.HasValue && currentPrice < above.Value)
            _lastAlertedAbove = false;

        if (below.HasValue && currentPrice <= below.Value && !_lastAlertedBelow)
        {
            _lastAlertedBelow = true;
            PlayAlertSoundLooping();
            MessageBox.Show(
                $"价格已跌至 {currentPrice:F2} {unitLabel}，低于设定值 {below.Value:F2}。",
                "价格提醒 - 低于设定值",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            StopAlertSound();
        }

        if (above.HasValue && currentPrice >= above.Value && !_lastAlertedAbove)
        {
            _lastAlertedAbove = true;
            PlayAlertSoundLooping();
            MessageBox.Show(
                $"价格已涨至 {currentPrice:F2} {unitLabel}，高于设定值 {above.Value:F2}。",
                "价格提醒 - 高于设定值",
                MessageBoxButton.OK,
                MessageBoxImage.Warning);
            StopAlertSound();
        }
    }

    private void Unit_Checked(object sender, RoutedEventArgs e)
    {
        // XAML 加载时可能先触发 Checked 再创建其它控件，需空判断避免 NullReferenceException
        if (RbCny == null) { _useCny = false; return; }
        _useCny = RbCny.IsChecked == true;
        // 切换单位时，应重新按新单位重置坐标轴范围，不能继续锁定旧横坐标
        if (ChkLockXAxis != null) ChkLockXAxis.IsChecked = false;
        _lockXAxis = false;
        if (GrpKline == null) return;
        GrpKline.Header = _useCny ? "价格折线图（最近 1 小时，元/克）" : "价格折线图（最近 1 小时，美元/盎司）";
        UpdateKline();
    }

    private void LockXAxis_Checked(object sender, RoutedEventArgs e)
    {
        _lockXAxis = ChkLockXAxis?.IsChecked == true;
    }

    private void UpdateForecastList(List<(TimeSpan Offset, double Price)> data)
    {
        LstForecast.Items.Clear();
        if (data.Count == 0)
        {
            LstForecast.Items.Add("暂无预测数据。");
            return;
        }

        var lastPrice = _history.Last().Price(_useCny);
        foreach (var (offset, price) in data)
        {
            var diff = price - lastPrice;
            var direction = diff > 0 ? "上涨" : diff < 0 ? "下跌" : "持平";
            LstForecast.Items.Add(
                $"未来 {offset.TotalMinutes,4:F0} 分钟: 预测 {price:F2} （{direction} {diff:+0.00;-0.00;0.00}）");
        }

        LstForecast.Items.Add("※ 仅供参考，不构成投资建议。");
    }
}

public record ForecastBandPoint(TimeSpan Offset, double Ma, double Upper, double Lower);

public record PricePoint(DateTime Time, double PriceUsd, double PriceCny)
{
    /// <summary> 按单位取价格，优先所选单位，若为 0 则用另一单位。 </summary>
    public double Price(bool useCny) =>
        useCny ? (PriceCny > 0 ? PriceCny : PriceUsd) : (PriceUsd > 0 ? PriceUsd : PriceCny);
}

public class GoldPriceDetail
{
    public string Name { get; set; } = "";
    public double YesterdayClose { get; set; }
    public double CurrentPrice { get; set; }
    public double HighPrice { get; set; }
    public double LowPrice { get; set; }
    public double OpenPrice { get; set; }
    public string Unit { get; set; } = "";
    public double ChangeAmount { get; set; }
    public double ChangePercent { get; set; }
}

public class GoldPriceResult
{
    public double? UsdPrice { get; set; }
    public double? CnyPrice { get; set; }
    public GoldPriceDetail? UsdDetails { get; set; }
    public GoldPriceDetail? CnyDetails { get; set; }

    public string Summary =>
        UsdPrice.HasValue && UsdDetails is not null
            ? $"现货黄金最新价: ${UsdPrice.Value:F2} 美元/盎司 ({UsdDetails.ChangePercent:+0.00;-0.00;0.00}%)"
            : "暂无汇总信息";
}

