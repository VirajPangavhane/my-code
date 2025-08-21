using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using SmartValveMatcherEngine;
using SmartValveMatcherEngine.Models;
using SmartValveMatcherEngine.Services;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;




namespace SmartValveMatcherEngine;

public class ValveDetector
{

    [CommandMethod("MATCH_AND_EXPORT")]
    public async void MatchAndExport()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;

        try
        {
            ed.WriteMessage("\nRunning valve matching...");
            MatchValves();

            await Task.Delay(3000);

            ed.WriteMessage("\nExporting matched valve blocks...");
            await ExportSmartPIDBlockDataAsync();
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n[ERROR] MATCH_AND_EXPORT failed: {ex.Message}");
        }
    }



    [CommandMethod("MATCH_VALVES")]
    public void MatchValves()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var db = doc.Database;
        var ed = doc.Editor;

        ed.WriteMessage("\n MATCH_VALVES started...");

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            // Step 2: Explode legacy valve blocks
            BlockExploder.ExplodeValveBlocks(ms, tr, ed);

            // Step 3: Fetch SmartMark metadata from API
            var apiMap = SmartMarkDataLoader.LoadSmartMarkMap(ed);

            // Step 4: Load pattern definitions
            string patternPath = @"C:\SmartValve\valve_patterns.json";
            List<ValvePattern> patterns;
            try
            {
                patterns = PatternLoader.LoadPatterns(patternPath);
                ed.WriteMessage($"\n Loaded {patterns.Count} valve pattern(s).");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Failed to load valve patterns: {ex.Message}");
                return;
            }

            // Step 5: Load valid tag prefixes from Excel
            string excelPath = @"C:\Valve_Detection\valve_prefixes.xlsx";
            List<string> validPrefixes;
            try
            {
                validPrefixes = TagPrefixLoader.LoadPrefixesFromExcel(excelPath);
                ed.WriteMessage($"\n Loaded {validPrefixes.Count} tag prefix(es) from Excel.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Failed to load tag prefixes: {ex.Message}");
                return;
            }

            string valveLayerExcelPath = @"C:\Valve_Detection\valve_layers.xlsx";
            List<string> allowedValveLayers;
            try
            {
                allowedValveLayers = TagPrefixLoader.LoadPrefixesFromExcel(valveLayerExcelPath)
                    .Select(l => l.Trim().ToUpper())
                    .Where(l => !string.IsNullOrWhiteSpace(l))
                    .ToList();
                ed.WriteMessage($"\n Loaded {allowedValveLayers.Count} valve layer(s) from Excel.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\n Failed to load valve layers: {ex.Message}");
                return;
            }

            // Step 6: Build regex for valid tags
            string pattern = @"^(" + string.Join("|", validPrefixes.Select(p => Regex.Escape(p))) + @")\d+$";
            Regex validTagRegex = new Regex(pattern, RegexOptions.IgnoreCase);

            // Step 7: Find AREA_ZONE polylines
            var areaZones = ms.Cast<ObjectId>()
                .Select(id => tr.GetObject(id, OpenMode.ForRead))
                .OfType<Polyline>()
                .Where(p => p.Layer == "AREA_ZONE" && p.Closed)
                .ToList();

            // Step 8: Find valve tags inside AREA_ZONEs
            var tagTexts = ms.Cast<ObjectId>()
                .Select(id => tr.GetObject(id, OpenMode.ForRead))
                .OfType<DBText>()
                .Where(txt =>
                {
                    var text = txt.TextString.Trim().ToUpper();
                    return !string.IsNullOrWhiteSpace(text)
                        && areaZones.Any(z => IsInside(txt.Position, z))
                        && validTagRegex.IsMatch(text);
                })
                .ToList();

            ed.WriteMessage($"\n Found {tagTexts.Count} valve tag(s).");

            // cache tag positions
            var allTagPositions = tagTexts.Select(t => (Tag: t, Pos: t.Position)).ToList();
            var allowedSet = new HashSet<string>(allowedValveLayers);

            int createdCount = 0;

            foreach (var tag in tagTexts)
            {
                ed.WriteMessage($"\n Analyzing tag: {tag.TextString}");

                // 🔹 Use cluster-ownership logic (only this tag’s entities)
                var cluster = GetOwnedClusterForTag(ms, tr, tag.Position, allTagPositions, allowedSet, 25.0, 2);

                if (cluster == null || cluster.Count == 0)
                {
                    ed.WriteMessage($"\n  No owned cluster found for tag: {tag.TextString}");
                    if (!IsSquareAlreadyMarked(tag.Position, ms, tr))
                    {
                        var squareSize = 7.0;
                        var half = squareSize / 2;
                        var lowerLeft = new Point2d(tag.Position.X - half, tag.Position.Y - half);

                        var square = new Polyline();
                        square.AddVertexAt(0, lowerLeft, 0, 0, 0);
                        square.AddVertexAt(1, new Point2d(lowerLeft.X + squareSize, lowerLeft.Y), 0, 0, 0);
                        square.AddVertexAt(2, new Point2d(lowerLeft.X + squareSize, lowerLeft.Y + squareSize), 0, 0, 0);
                        square.AddVertexAt(3, new Point2d(lowerLeft.X, lowerLeft.Y + squareSize), 0, 0, 0);
                        square.Closed = true;
                        square.ColorIndex = 1;

                        ms.AppendEntity(square);
                        tr.AddNewlyCreatedDBObject(square, true);
                    }
                    continue;
                }

                ed.WriteMessage($"\n  Using cluster (entities: {cluster.Count})");

                var lines = cluster.OfType<Line>().Where(l => l.Length < 30).ToList();
                var circles = cluster.OfType<Circle>().ToList();
                var arcs = cluster.OfType<Arc>().ToList();
                var pls = cluster.OfType<Polyline>().Where(p => !string.Equals(p.Layer, "AREA_ZONE", StringComparison.OrdinalIgnoreCase)).ToList();
                var solids = cluster.OfType<Solid>().ToList();
                var hatches = cluster.OfType<Hatch>().ToList();
                var extras = new List<DBText>(); // don’t pull other valve tags

                string valveType = PatternMatcher.MatchGeometryToPattern(lines, circles, arcs, pls, solids, hatches, patterns);

                if (string.IsNullOrWhiteSpace(valveType))
                {
                    ed.WriteMessage($"  Unmatched valve at tag: {tag.TextString}");
                    // draw red square...
                    continue;
                }

                ed.WriteMessage($"  Matched as: {valveType}");
                RemoveRedSquareIfExists(tag.Position, ms, tr);

                var zone = areaZones.FirstOrDefault(z => IsInside(tag.Position, z));
                var (facility, subFacility) = GetSmartMarkData(zone);
                var key = (facility, subFacility);

                Dictionary<string, string> extraAttrs = apiMap.ContainsKey(key) ? apiMap[key] : new();
                var hardcodedAttrs = ValveTypeDataProvider.GetAttributesForValve(valveType);
                var mergedAttrs = new Dictionary<string, string>(hardcodedAttrs);
                foreach (var kvp in extraAttrs)
                    mergedAttrs[kvp.Key] = kvp.Value;

                BlockBuilder.CreateValveBlock(
                    bt, ms, tr,
                    tag, tag.TextString.Trim(), valveType,
                    lines, circles, arcs, solids, hatches, pls, extras,
                    mergedAttrs,
                    ref createdCount
                );
            }

            ed.WriteMessage($"\n Rebuilt {createdCount} valve block(s).");
            tr.Commit();
        }

        ed.WriteMessage("\n Valve analysis complete.");
    }


    bool IsSquareAlreadyMarked(Point3d position, BlockTableRecord ms, Transaction tr, double tolerance = 1.0)
    {
        foreach (ObjectId id in ms)
        {
            var obj = tr.GetObject(id, OpenMode.ForRead);
            if (obj is Polyline pl &&
                pl.Closed &&
                pl.NumberOfVertices == 4 &&
                pl.ColorIndex == 1) // Red
            {
                var plCenter = GetPolylineCenter(pl);
                if (plCenter.DistanceTo(position) < tolerance)
                    return true;
            }
        }
        return false;
    }





    private Point3d GetPolylineCenter(Polyline pl)
    {
        var bounds = pl.GeometricExtents;
        return new Point3d(
            (bounds.MinPoint.X + bounds.MaxPoint.X) / 2,
            (bounds.MinPoint.Y + bounds.MaxPoint.Y) / 2,
            0
        );
    }

    private void RemoveRedSquareIfExists(Point3d position, BlockTableRecord ms, Transaction tr, double tolerance = 1.0)
    {
        foreach (ObjectId id in ms)
        {
            var obj = tr.GetObject(id, OpenMode.ForRead);
            if (obj is Polyline pl && pl.Closed && pl.NumberOfVertices == 4 && pl.ColorIndex == 1)
            {
                var center = GetPolylineCenter(pl);
                if (center.DistanceTo(position) < tolerance)
                {
                    pl.UpgradeOpen();
                    pl.Erase();
                    break;
                }
            }
        }
    }

    private static List<Entity> GetOwnedClusterForTag(
    BlockTableRecord ms,
    Transaction tr,
    Point3d tagPos,
    List<(DBText Tag, Point3d Pos)> allTags,
    HashSet<string> allowedLayerSet,
    double proximity,
    double connectTolerance)
    {
        var candidates = new List<Entity>();
        foreach (ObjectId id in ms)
        {
            if (tr.GetObject(id, OpenMode.ForRead) is not Entity ent) continue;
            if (ent.IsErased || ent.IsDisposed || !ent.Bounds.HasValue) continue;
            if (!(ent is Line || ent is Circle || ent is Arc || ent is Polyline || ent is Solid || ent is Hatch)) continue;

            var layer = (ent.Layer ?? "").Trim().ToUpperInvariant();
            if (!allowedLayerSet.Contains(layer)) continue;

            var center = GetCenter(ent.GeometricExtents);
            if (center.DistanceTo(tagPos) <= proximity)
                candidates.Add(ent);
        }

        if (candidates.Count == 0) return new();

        var clusters = BuildClusters(candidates, connectTolerance);

        List<Entity>? owned = null;
        double bestDist = double.MaxValue;

        foreach (var cluster in clusters)
        {
            var centroid = ClusterCentroid(cluster);

            // find nearest tag
            double minDist = double.MaxValue;
            Point3d nearestTagPos = Point3d.Origin;
            foreach (var t in allTags)
            {
                double d = t.Pos.DistanceTo(centroid);
                if (d < minDist) { minDist = d; nearestTagPos = t.Pos; }
            }

            // only keep if this tag is the nearest
            double myDist = tagPos.DistanceTo(centroid);
            if (Math.Abs(myDist - minDist) < 10) // allow small tolerance
            {
                if (myDist < bestDist)
                {
                    bestDist = myDist;
                    owned = cluster;
                }
            }
        }

        return owned ?? new();
    }

    private static List<List<Entity>> BuildClusters(List<Entity> entities, double tol)
    {
        var clusters = new List<List<Entity>>();
        var visited = new HashSet<Entity>();

        foreach (var e in entities)
        {
            if (visited.Contains(e)) continue;
            var q = new Queue<Entity>();
            var cluster = new List<Entity>();
            q.Enqueue(e);
            visited.Add(e);

            while (q.Count > 0)
            {
                var cur = q.Dequeue();
                cluster.Add(cur);

                foreach (var other in entities)
                {
                    if (visited.Contains(other)) continue;
                    if (EntitiesAreClose(cur, other, tol))
                    {
                        visited.Add(other);
                        q.Enqueue(other);
                    }
                }
            }

            clusters.Add(cluster);
        }
        return clusters;
    }

    private static bool EntitiesAreClose(Entity a, Entity b, double tol)
    {
        try
        {
            var A = a.GeometricExtents; var B = b.GeometricExtents;
            double dx = Math.Max(0, Math.Max(A.MinPoint.X - B.MaxPoint.X, B.MinPoint.X - A.MaxPoint.X));
            double dy = Math.Max(0, Math.Max(A.MinPoint.Y - B.MaxPoint.Y, B.MinPoint.Y - A.MaxPoint.Y));
            return Math.Sqrt(dx * dx + dy * dy) <= tol;
        }
        catch { return false; }
    }

    private static Point3d ClusterCentroid(List<Entity> cluster)
    {
        double sx = 0, sy = 0; int n = 0;
        foreach (var e in cluster)
        {
            if (!e.Bounds.HasValue) continue;
            var c = GetCenter(e.GeometricExtents);
            sx += c.X; sy += c.Y; n++;
        }
        return n == 0 ? Point3d.Origin : new Point3d(sx / n, sy / n, 0);
    }

    private static Point3d GetCenter(Extents3d ext)
    {
        return new Point3d(
            (ext.MinPoint.X + ext.MaxPoint.X) / 2.0,
            (ext.MinPoint.Y + ext.MaxPoint.Y) / 2.0,
            0
        );
    }
























    [CommandMethod("SMARTPID_VALVE_EXPORTBLOCKS")]
    public async void ExportSmartPIDBlockData()
    {
        await ExportSmartPIDBlockDataAsync();
    }

    public async Task ExportSmartPIDBlockDataAsync()
    {
        var doc = Application.DocumentManager.MdiActiveDocument;
        var ed = doc.Editor;
        var db = doc.Database;

        var blockDataList = new List<Dictionary<string, string>>();

        using (var tr = db.TransactionManager.StartTransaction())
        {
            var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            var ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objId in ms)
            {
                if (tr.GetObject(objId, OpenMode.ForRead) is BlockReference br)
                {
                    if (br.AttributeCollection.Count == 0) continue;

                    bool isSmartPIDBlock = false;
                    var data = new Dictionary<string, string>
                    {
                        { "BlockName", br.Name }
                    };

                    foreach (ObjectId attId in br.AttributeCollection)
                    {
                        if (tr.GetObject(attId, OpenMode.ForRead) is AttributeReference attRef)
                        {
                            string tag = attRef.Tag.Trim();
                            string value = attRef.TextString.Trim();
                            data[tag] = value;

                            if (tag.Equals("VALVE_TAG", StringComparison.OrdinalIgnoreCase))
                                isSmartPIDBlock = true;
                        }
                    }

                    if (isSmartPIDBlock)
                        blockDataList.Add(data);
                }
            }

            tr.Commit();
        }

        try
        {
            using var client = new HttpClient();
            client.DefaultRequestHeaders.Add("Authorization", "token cbe182e6b5c60f6:8d7e0ae7bad6198");

            string jobId = await FetchCurrentJobId(client, ed);
            if (string.IsNullOrWhiteSpace(jobId))
            {
                ed.WriteMessage("\n[ERROR] No job ID found to update.");
                return;
            }

            string blockJson = JsonConvert.SerializeObject(blockDataList, Formatting.Indented);

            var updatePayload = new
            {
                valve_output_data = blockJson
            };

            var content = new StringContent(JsonConvert.SerializeObject(updatePayload), Encoding.UTF8, "application/json");
            var updateResponse = await client.PatchAsync(
                $"https://enaibotdevfrappe.inventivebizsol.co.in/api/v2/document/Instrumentation Files/guur80oug4",
                content
            );

            var responseText = await updateResponse.Content.ReadAsStringAsync();
            ed.WriteMessage($"\n[Frappe Response] {responseText}");

            if (updateResponse.IsSuccessStatusCode)
            {
                ed.WriteMessage($"\nExported {blockDataList.Count} SmartPID valve block(s) successfully.");
            }
            else
            {
                ed.WriteMessage($"\n[ERROR] Failed to push data to API:\n{responseText}");
            }
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n[ERROR] Export failed: {ex.Message}");
        }
    }

    private async Task<string> FetchCurrentJobId(HttpClient client, Editor ed)
    {
        try
        {
            var request = new HttpRequestMessage(
                HttpMethod.Get,
                "https://enaibotdevfrappe.inventivebizsol.co.in/api/v2/document/Instrumentation Files?fields=[\"*\"]&filters=[[\"status\", \"=\", \"IN_PROCESS\"]]"
            );
            request.Headers.Add("Authorization", "token cbe182e6b5c60f6:8d7e0ae7bad6198");

            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();

            var jsonResponse = JsonDocument.Parse(await response.Content.ReadAsStringAsync());
            var dataArray = jsonResponse.RootElement.GetProperty("data");

            if (dataArray.GetArrayLength() == 0)
            {
                ed.WriteMessage("\n[ERROR] No IN_PROCESS job found.");
                return "";
            }

            return dataArray[0].GetProperty("name").GetString();
        }
        catch (System.Exception ex)
        {
            ed.WriteMessage($"\n[ERROR] Failed to fetch job ID: {ex.Message}");
            return "";
        }
    }

    private bool IsInside(Point3d pt, Polyline pl)
    {
        try
        {
            var ext = pl.GeometricExtents;
            return pt.X >= ext.MinPoint.X && pt.X <= ext.MaxPoint.X &&
                   pt.Y >= ext.MinPoint.Y && pt.Y <= ext.MaxPoint.Y;
        }
        catch { return false; }
    }

    private (string Facility, string SubFacility) GetSmartMarkData(Polyline zone)
    {
        string group = "UNKNOWN";
        string subgroup = "UNKNOWN";

        try
        {
            var rb = zone.GetXDataForApplication("SMARTMARK");
            if (rb != null)
            {
                var data = rb.AsArray();
                if (data.Length >= 3 && data[0].TypeCode == 1001 && data[0].Value.ToString() == "SMARTMARK")
                {
                    if (data[1].TypeCode == 1000) group = data[1].Value.ToString();
                    if (data[2].TypeCode == 1000) subgroup = data[2].Value.ToString();
                }
            }
        }
        catch { }

        return (group.Trim(), subgroup.Trim());
    }
}
