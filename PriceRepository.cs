using System;
using System.Collections.Generic;
using Microsoft.Data.Sqlite;

namespace GoldDesktop;

public class PriceRepository
{
    private readonly string _connectionString;

    public PriceRepository(string dbPath)
    {
        _connectionString = $"Data Source={dbPath}";
        Init();
    }

    private void Init()
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText =
            """
            CREATE TABLE IF NOT EXISTS ticks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                time_utc TEXT NOT NULL,
                price_usd REAL,
                price_cny REAL
            );
            """;
        cmd.ExecuteNonQuery();
    }

    public void ClearAll()
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText = "DELETE FROM ticks;";
        cmd.ExecuteNonQuery();
    }

    public void SaveTick(DateTime timeUtc, double? priceUsd, double? priceCny)
    {
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText =
            "INSERT INTO ticks (time_utc, price_usd, price_cny) VALUES ($t, $u, $c);";
        cmd.Parameters.AddWithValue("$t", timeUtc.ToString("O"));
        cmd.Parameters.AddWithValue("$u", (object?)priceUsd ?? DBNull.Value);
        cmd.Parameters.AddWithValue("$c", (object?)priceCny ?? DBNull.Value);
        cmd.ExecuteNonQuery();
    }

    public List<Candle> GetRecentCandles(TimeSpan window, TimeSpan candleSize)
    {
        var nowUtc = DateTime.UtcNow;
        var fromUtc = nowUtc - window;
        using var conn = new SqliteConnection(_connectionString);
        conn.Open();
        var cmd = conn.CreateCommand();
        cmd.CommandText =
            "SELECT time_utc, price_usd, price_cny FROM ticks WHERE time_utc >= $from AND (price_usd IS NOT NULL OR price_cny IS NOT NULL) ORDER BY time_utc;";
        cmd.Parameters.AddWithValue("$from", fromUtc.ToString("O"));
        using var reader = cmd.ExecuteReader();

        var points = new List<PricePoint>();
        while (reader.Read())
        {
            var t = DateTime.Parse(reader.GetString(0), null, System.Globalization.DateTimeStyles.RoundtripKind);
            var u = reader.IsDBNull(1) ? 0.0 : reader.GetDouble(1);
            var c = reader.IsDBNull(2) ? 0.0 : reader.GetDouble(2);
            points.Add(new PricePoint(t, u, c));
        }

        var candles = new List<Candle>();
        if (points.Count == 0) return candles;

        // 聚合成 K 线时使用美元价，若无则用人民币价
        double GetPrice(PricePoint pt) => pt.PriceUsd > 0 ? pt.PriceUsd : pt.PriceCny;
        var bucketStart = AlignTime(points[0].Time, candleSize);
        double open = GetPrice(points[0]);
        double high = open;
        double low = open;
        double close = open;

        foreach (var pt in points)
        {
            var bucket = AlignTime(pt.Time, candleSize);
            var p = GetPrice(pt);
            if (bucket != bucketStart)
            {
                candles.Add(new Candle(bucketStart, open, high, low, close));
                bucketStart = bucket;
                open = high = low = close = p;
            }
            else
            {
                if (p > high) high = p;
                if (p < low) low = p;
                close = p;
            }
        }

        candles.Add(new Candle(bucketStart, open, high, low, close));
        return candles;
    }

    private static DateTime AlignTime(DateTime t, TimeSpan step)
    {
        var ticks = t.Ticks / step.Ticks * step.Ticks;
        return new DateTime(ticks, DateTimeKind.Utc);
    }
}

public record Candle(DateTime StartTimeUtc, double Open, double High, double Low, double Close);

