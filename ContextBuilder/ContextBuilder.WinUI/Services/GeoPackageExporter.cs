using System.Globalization;
using System.Text;
using ContextBuilder.WinUI.Models;
using Microsoft.Data.Sqlite;

namespace ContextBuilder.WinUI.Services;

public sealed class GeoPackageExporter
{
    private const int GpkgApplicationId = 0x47504B47; // "GPKG"
    private const int GpkgUserVersion = 10300; // GeoPackage 1.3

    public string Export(
        string folderPath,
        string sourceEpsg,
        string targetEpsg,
        IReadOnlyCollection<LayerResult> layers,
        IReadOnlyList<(GeoPoint Point, double Elevation)>? elevationGrid)
    {
        Directory.CreateDirectory(folderPath);
        var outputPath = Path.Combine(folderPath, "context_layers.gpkg");
        if (File.Exists(outputPath))
        {
            File.Delete(outputPath);
        }

        var srsId = ParseSrsId(targetEpsg);
        using var conn = new SqliteConnection($"Data Source={outputPath}");
        conn.Open();

        using (var cmd = conn.CreateCommand())
        {
            cmd.CommandText = $"""
                PRAGMA application_id={GpkgApplicationId};
                PRAGMA user_version={GpkgUserVersion};
                """;
            cmd.ExecuteNonQuery();
        }

        CreateMetadataTables(conn);
        SeedSpatialRefs(conn, srsId, targetEpsg);

        foreach (var layer in layers)
        {
            WriteLayer(conn, layer, sourceEpsg, targetEpsg, srsId);
        }

        if (elevationGrid is { Count: > 0 })
        {
            WriteElevationLayer(conn, elevationGrid, sourceEpsg, targetEpsg, srsId);
        }

        return outputPath;
    }

    private static void CreateMetadataTables(SqliteConnection conn)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            CREATE TABLE gpkg_spatial_ref_sys (
                srs_name TEXT NOT NULL,
                srs_id INTEGER NOT NULL PRIMARY KEY,
                organization TEXT NOT NULL,
                organization_coordsys_id INTEGER NOT NULL,
                definition TEXT NOT NULL,
                description TEXT
            );

            CREATE TABLE gpkg_contents (
                table_name TEXT NOT NULL PRIMARY KEY,
                data_type TEXT NOT NULL,
                identifier TEXT UNIQUE,
                description TEXT DEFAULT '',
                last_change DATETIME NOT NULL DEFAULT (strftime('%Y-%m-%dT%H:%M:%fZ','now')),
                min_x DOUBLE,
                min_y DOUBLE,
                max_x DOUBLE,
                max_y DOUBLE,
                srs_id INTEGER,
                CONSTRAINT fk_gc_r_srs_id FOREIGN KEY (srs_id) REFERENCES gpkg_spatial_ref_sys(srs_id)
            );

            CREATE TABLE gpkg_geometry_columns (
                table_name TEXT NOT NULL,
                column_name TEXT NOT NULL,
                geometry_type_name TEXT NOT NULL,
                srs_id INTEGER NOT NULL,
                z TINYINT NOT NULL,
                m TINYINT NOT NULL,
                PRIMARY KEY (table_name, column_name),
                CONSTRAINT fk_gc_tn FOREIGN KEY (table_name) REFERENCES gpkg_contents(table_name),
                CONSTRAINT fk_gc_srs FOREIGN KEY (srs_id) REFERENCES gpkg_spatial_ref_sys(srs_id)
            );
            """;
        cmd.ExecuteNonQuery();
    }

    private static void SeedSpatialRefs(SqliteConnection conn, int srsId, string targetEpsg)
    {
        using var cmd = conn.CreateCommand();
        cmd.CommandText = """
            INSERT INTO gpkg_spatial_ref_sys(srs_name,srs_id,organization,organization_coordsys_id,definition,description) VALUES
            ('Undefined Cartesian',-1,'NONE',-1,'undefined','undefined cartesian coordinate reference system'),
            ('Undefined Geographic',0,'NONE',0,'undefined','undefined geographic coordinate reference system'),
            ('WGS 84 geodetic',4326,'EPSG',4326,'GEOGCS["WGS 84",DATUM["WGS_1984",SPHEROID["WGS 84",6378137,298.257223563]],PRIMEM["Greenwich",0],UNIT["degree",0.0174532925199433]]',''),
            ('WGS 84 / Pseudo-Mercator',3857,'EPSG',3857,'PROJCS["WGS 84 / Pseudo-Mercator",GEOGCS["WGS 84",DATUM["WGS_1984",SPHEROID["WGS 84",6378137,298.257223563]],PRIMEM["Greenwich",0],UNIT["degree",0.0174532925199433]],PROJECTION["Mercator_1SP"],PARAMETER["central_meridian",0],PARAMETER["scale_factor",1],PARAMETER["false_easting",0],PARAMETER["false_northing",0],UNIT["metre",1]]','');
            """;
        cmd.ExecuteNonQuery();

        if (srsId is 4326 or 3857)
        {
            return;
        }

        using var insertCustom = conn.CreateCommand();
        insertCustom.CommandText = """
            INSERT OR IGNORE INTO gpkg_spatial_ref_sys
            (srs_name,srs_id,organization,organization_coordsys_id,definition,description)
            VALUES ($name,$srs,'EPSG',$srs,'undefined','User-defined EPSG from ContextBuilder');
            """;
        insertCustom.Parameters.AddWithValue("$name", targetEpsg);
        insertCustom.Parameters.AddWithValue("$srs", srsId);
        insertCustom.ExecuteNonQuery();
    }

    private static void WriteLayer(
        SqliteConnection conn,
        LayerResult layer,
        string sourceEpsg,
        string targetEpsg,
        int srsId)
    {
        var tableName = layer.Layer.ToString().ToLowerInvariant();
        var withHeight = layer.Layer == ContextLayer.Buildings;
        CreateLayerTable(conn, tableName, withHeight: withHeight);

        var features = new List<(byte[] Blob, double MinX, double MinY, double MaxX, double MaxY, double? BaseHeightM, double? HeightM, string? HeightSource)>();
        foreach (var line in layer.LineStrings)
        {
            var coords = line.Select(p => GeoProjection.ProjectForTarget(p, sourceEpsg, targetEpsg)).ToList();
            if (coords.Count < 2)
            {
                continue;
            }

            var built = BuildGeomBlob(2, coords, srsId);
            features.Add((built.Blob, built.MinX, built.MinY, built.MaxX, built.MaxY, null, null, null));
        }

        for (var i = 0; i < layer.Polygons.Count; i++)
        {
            var polygon = layer.Polygons[i];
            var coords = polygon.Select(p => GeoProjection.ProjectForTarget(p, sourceEpsg, targetEpsg)).ToList();
            if (coords.Count < 3)
            {
                continue;
            }

            if (!SamePoint(coords[0], coords[^1]))
            {
                coords.Add(coords[0]);
            }

            var built = BuildGeomBlob(3, coords, srsId);
            double? bh = null;
            double? h = null;
            string? hs = null;
            if (withHeight && i < layer.BuildingFootprints.Count)
            {
                bh = Math.Round(layer.BuildingFootprints[i].BaseHeightMeters, 3);
                h = Math.Round(layer.BuildingFootprints[i].HeightMeters, 3);
                hs = layer.BuildingFootprints[i].HeightSource;
            }

            features.Add((built.Blob, built.MinX, built.MinY, built.MaxX, built.MaxY, bh, h, hs));
        }

        InsertFeatures(conn, tableName, features, withHeight: withHeight);
        RegisterLayer(
            conn,
            tableName,
            srsId,
            features.Select(x => (x.Blob, x.MinX, x.MinY, x.MaxX, x.MaxY)).ToList(),
            "GEOMETRY");
    }

    private static void WriteElevationLayer(
        SqliteConnection conn,
        IReadOnlyList<(GeoPoint Point, double Elevation)> elevationGrid,
        string sourceEpsg,
        string targetEpsg,
        int srsId)
    {
        const string tableName = "lidar_toposurface_points";
        CreateLayerTable(conn, tableName, withElevation: true);

        var features = new List<(byte[] Blob, double MinX, double MinY, double MaxX, double MaxY, double Elevation)>();
        foreach (var sample in elevationGrid)
        {
            var p = GeoProjection.ProjectForTarget(sample.Point, sourceEpsg, targetEpsg);
            var blob = BuildPointBlob(p.X, p.Y, srsId);
            features.Add((blob, p.X, p.Y, p.X, p.Y, sample.Elevation));
        }

        using (var tx = conn.BeginTransaction())
        {
            using var cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = $"INSERT INTO {tableName}(geom,elevation) VALUES ($geom,$elevation);";
            var geomParam = cmd.CreateParameter();
            geomParam.ParameterName = "$geom";
            cmd.Parameters.Add(geomParam);
            var elevationParam = cmd.CreateParameter();
            elevationParam.ParameterName = "$elevation";
            cmd.Parameters.Add(elevationParam);

            foreach (var f in features)
            {
                geomParam.Value = f.Blob;
                elevationParam.Value = f.Elevation;
                cmd.ExecuteNonQuery();
            }

            tx.Commit();
        }

        RegisterLayer(
            conn,
            tableName,
            srsId,
            features.Select(x => (x.Blob, x.MinX, x.MinY, x.MaxX, x.MaxY)).ToList(),
            "POINT");
    }

    private static void CreateLayerTable(SqliteConnection conn, string tableName, bool withElevation = false, bool withHeight = false)
    {
        using var cmd = conn.CreateCommand();
        if (withElevation)
        {
            cmd.CommandText = $"CREATE TABLE {tableName}(id INTEGER PRIMARY KEY AUTOINCREMENT, geom BLOB NOT NULL, elevation REAL);";
        }
        else if (withHeight)
        {
            cmd.CommandText = $"CREATE TABLE {tableName}(id INTEGER PRIMARY KEY AUTOINCREMENT, geom BLOB NOT NULL, base_height_m REAL, height_m REAL, height_source TEXT);";
        }
        else
        {
            cmd.CommandText = $"CREATE TABLE {tableName}(id INTEGER PRIMARY KEY AUTOINCREMENT, geom BLOB NOT NULL);";
        }
        cmd.ExecuteNonQuery();
    }

    private static void InsertFeatures(
        SqliteConnection conn,
        string tableName,
        IReadOnlyList<(byte[] Blob, double MinX, double MinY, double MaxX, double MaxY, double? BaseHeightM, double? HeightM, string? HeightSource)> features,
        bool withHeight = false)
    {
        using var tx = conn.BeginTransaction();
        using var cmd = conn.CreateCommand();
        cmd.Transaction = tx;
        cmd.CommandText = withHeight
            ? $"INSERT INTO {tableName}(geom,base_height_m,height_m,height_source) VALUES ($geom,$baseHeight,$height,$heightSource);"
            : $"INSERT INTO {tableName}(geom) VALUES ($geom);";
        var geomParam = cmd.CreateParameter();
        geomParam.ParameterName = "$geom";
        cmd.Parameters.Add(geomParam);
        SqliteParameter? baseHeightParam = null;
        SqliteParameter? heightParam = null;
        SqliteParameter? heightSourceParam = null;
        if (withHeight)
        {
            baseHeightParam = cmd.CreateParameter();
            baseHeightParam.ParameterName = "$baseHeight";
            cmd.Parameters.Add(baseHeightParam);
            heightParam = cmd.CreateParameter();
            heightParam.ParameterName = "$height";
            cmd.Parameters.Add(heightParam);
            heightSourceParam = cmd.CreateParameter();
            heightSourceParam.ParameterName = "$heightSource";
            cmd.Parameters.Add(heightSourceParam);
        }
        foreach (var f in features)
        {
            geomParam.Value = f.Blob;
            if (withHeight && heightParam is not null)
            {
                if (baseHeightParam is not null)
                {
                    baseHeightParam.Value = (object?)f.BaseHeightM ?? DBNull.Value;
                }
                heightParam.Value = (object?)f.HeightM ?? DBNull.Value;
                if (heightSourceParam is not null)
                {
                    heightSourceParam.Value = (object?)f.HeightSource ?? DBNull.Value;
                }
            }

            cmd.ExecuteNonQuery();
        }

        tx.Commit();
    }

    private static void RegisterLayer(
        SqliteConnection conn,
        string tableName,
        int srsId,
        IReadOnlyList<(byte[] Blob, double MinX, double MinY, double MaxX, double MaxY)> features,
        string geometryType)
    {
        double? minX = null;
        double? minY = null;
        double? maxX = null;
        double? maxY = null;
        foreach (var f in features)
        {
            minX = !minX.HasValue ? f.MinX : Math.Min(minX.Value, f.MinX);
            minY = !minY.HasValue ? f.MinY : Math.Min(minY.Value, f.MinY);
            maxX = !maxX.HasValue ? f.MaxX : Math.Max(maxX.Value, f.MaxX);
            maxY = !maxY.HasValue ? f.MaxY : Math.Max(maxY.Value, f.MaxY);
        }

        using (var contentsCmd = conn.CreateCommand())
        {
            contentsCmd.CommandText = """
                INSERT INTO gpkg_contents(table_name,data_type,identifier,description,min_x,min_y,max_x,max_y,srs_id)
                VALUES ($table,'features',$id,'', $minX,$minY,$maxX,$maxY,$srs);
                """;
            contentsCmd.Parameters.AddWithValue("$table", tableName);
            contentsCmd.Parameters.AddWithValue("$id", tableName);
            contentsCmd.Parameters.AddWithValue("$minX", (object?)minX ?? DBNull.Value);
            contentsCmd.Parameters.AddWithValue("$minY", (object?)minY ?? DBNull.Value);
            contentsCmd.Parameters.AddWithValue("$maxX", (object?)maxX ?? DBNull.Value);
            contentsCmd.Parameters.AddWithValue("$maxY", (object?)maxY ?? DBNull.Value);
            contentsCmd.Parameters.AddWithValue("$srs", srsId);
            contentsCmd.ExecuteNonQuery();
        }

        using var geomColsCmd = conn.CreateCommand();
        geomColsCmd.CommandText = """
            INSERT INTO gpkg_geometry_columns(table_name,column_name,geometry_type_name,srs_id,z,m)
            VALUES ($table,'geom',$type,$srs,0,0);
            """;
        geomColsCmd.Parameters.AddWithValue("$table", tableName);
        geomColsCmd.Parameters.AddWithValue("$type", geometryType);
        geomColsCmd.Parameters.AddWithValue("$srs", srsId);
        geomColsCmd.ExecuteNonQuery();
    }

    private static (byte[] Blob, double MinX, double MinY, double MaxX, double MaxY) BuildGeomBlob(
        int wkbType,
        IReadOnlyList<(double X, double Y)> coords,
        int srsId)
    {
        var minX = coords.Min(c => c.X);
        var minY = coords.Min(c => c.Y);
        var maxX = coords.Max(c => c.X);
        var maxY = coords.Max(c => c.Y);

        var wkb = new List<byte>();
        wkb.Add(1); // little endian
        wkb.AddRange(BitConverter.GetBytes(wkbType));
        if (wkbType == 2)
        {
            wkb.AddRange(BitConverter.GetBytes(coords.Count));
            foreach (var c in coords)
            {
                wkb.AddRange(BitConverter.GetBytes(c.X));
                wkb.AddRange(BitConverter.GetBytes(c.Y));
            }
        }
        else
        {
            // polygon single ring
            wkb.AddRange(BitConverter.GetBytes(1));
            wkb.AddRange(BitConverter.GetBytes(coords.Count));
            foreach (var c in coords)
            {
                wkb.AddRange(BitConverter.GetBytes(c.X));
                wkb.AddRange(BitConverter.GetBytes(c.Y));
            }
        }

        var blob = BuildGpkgGeometryBlob(srsId, minX, minY, maxX, maxY, wkb.ToArray());
        return (blob, minX, minY, maxX, maxY);
    }

    private static byte[] BuildPointBlob(double x, double y, int srsId)
    {
        var wkb = new List<byte>();
        wkb.Add(1);
        wkb.AddRange(BitConverter.GetBytes(1)); // Point
        wkb.AddRange(BitConverter.GetBytes(x));
        wkb.AddRange(BitConverter.GetBytes(y));
        return BuildGpkgGeometryBlob(srsId, x, y, x, y, wkb.ToArray());
    }

    private static byte[] BuildGpkgGeometryBlob(
        int srsId,
        double minX,
        double minY,
        double maxX,
        double maxY,
        byte[] wkb)
    {
        var blob = new List<byte>(8 + 32 + wkb.Length)
        {
            0x47, 0x50, // GP
            0, // version
            0x03 // flags: little-endian + envelope XY
        };
        blob.AddRange(BitConverter.GetBytes(srsId));
        blob.AddRange(BitConverter.GetBytes(minX));
        blob.AddRange(BitConverter.GetBytes(maxX));
        blob.AddRange(BitConverter.GetBytes(minY));
        blob.AddRange(BitConverter.GetBytes(maxY));
        blob.AddRange(wkb);
        return blob.ToArray();
    }

    private static bool SamePoint((double X, double Y) a, (double X, double Y) b)
    {
        return Math.Abs(a.X - b.X) < 1e-7 && Math.Abs(a.Y - b.Y) < 1e-7;
    }

    private static int ParseSrsId(string epsg)
    {
        var digits = new string(epsg.Where(char.IsDigit).ToArray());
        if (int.TryParse(digits, NumberStyles.Integer, CultureInfo.InvariantCulture, out var id))
        {
            return id;
        }

        return 4326;
    }
}
