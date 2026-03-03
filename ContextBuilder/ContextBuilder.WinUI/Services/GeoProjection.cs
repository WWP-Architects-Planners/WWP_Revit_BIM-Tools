using ContextBuilder.WinUI.Models;

namespace ContextBuilder.WinUI.Services;

public static class GeoProjection
{
    private const double EarthRadiusMeters = 6378137.0;

    public static (double X, double Y) ToLocalMeters(GeoPoint origin, GeoPoint point)
    {
        var latRad = DegreesToRadians(point.Latitude);
        var lonRad = DegreesToRadians(point.Longitude);
        var originLatRad = DegreesToRadians(origin.Latitude);
        var originLonRad = DegreesToRadians(origin.Longitude);

        var x = (lonRad - originLonRad) * EarthRadiusMeters * Math.Cos(originLatRad);
        var y = (latRad - originLatRad) * EarthRadiusMeters;
        return (x, y);
    }

    public static GeoPoint OffsetByMeters(GeoPoint origin, double dxMeters, double dyMeters)
    {
        var dLat = (dyMeters / EarthRadiusMeters) * (180.0 / Math.PI);
        var dLon = (dxMeters / (EarthRadiusMeters * Math.Cos(DegreesToRadians(origin.Latitude)))) * (180.0 / Math.PI);
        return new GeoPoint(origin.Latitude + dLat, origin.Longitude + dLon);
    }

    public static (double X, double Y) ToWebMercator(GeoPoint point)
    {
        var x = EarthRadiusMeters * DegreesToRadians(point.Longitude);
        var y = EarthRadiusMeters * Math.Log(Math.Tan(Math.PI / 4.0 + DegreesToRadians(point.Latitude) / 2.0));
        return (x, y);
    }

    public static (double X, double Y) ProjectForTarget(GeoPoint point, string sourceEpsg, string targetEpsg)
    {
        // Current data providers return WGS84 lat/lon coordinates.
        // If a non-4326 source is selected in UI, treat it as metadata for now.
        _ = sourceEpsg;
        return targetEpsg.Contains("3857", StringComparison.OrdinalIgnoreCase)
            ? ToWebMercator(point)
            : (point.Longitude, point.Latitude);
    }

    public static bool IsTargetSupported(string targetEpsg)
    {
        return targetEpsg.Contains("4326", StringComparison.OrdinalIgnoreCase) ||
               targetEpsg.Contains("3857", StringComparison.OrdinalIgnoreCase);
    }

    private static double DegreesToRadians(double deg) => deg * Math.PI / 180.0;
}
