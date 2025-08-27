using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;                 // <-- FIXED (for File.ReadLines, Path, etc.)
using System.Text.RegularExpressions;
using ClosedXML.Excel;
using FirstProject.Models;       // <-- FIXED (your models are in FirstProject.Models)

using FirstProject.Models;

namespace FirstProject.Helper
{
    public static class Techstudioconverter
    {
        private const string TS = "yyyy-MM-dd HH:mm:ss,fff";

        private static readonly Regex rxVersion = new(@" - (?<ver>\d+\.\d+\.\d+\.\d+) - ", RegexOptions.Compiled);
        private static readonly Regex rxAdded = new(
            @"^(?<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3}).*? - Added command: (?<cmd>[A-Za-z0-9]+) Id: (?<id>\d+)",
            RegexOptions.Compiled);
        private static readonly Regex rxMsgIdResp = new(
            @"^(?<ts>\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2},\d{3}).*? - Response message Id: (?<id>\d+)",
            RegexOptions.Compiled);

        private static readonly Regex rxNotifyEndpoint = new(@" - Notify to UI for EndPoint - (?<id>[0-9A-F]+) - Mesh", RegexOptions.Compiled);
        private static readonly Regex rxEndpointOk = new(@"Endpoint connection successful, id: .*?\[(?<id>[0-9A-F]+)\]", RegexOptions.Compiled);

        private static readonly Regex rxSuccessLocal = new(@"Local radio connected successfully", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex rxSuccess8052 = new(@"80 52 completed successfully", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex rxSuccessMfg = new(@"Manufacturing command processed Successfully", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        private static readonly Regex rxFailedConnect = new(@"Failed to connect the endpoint", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        private static readonly Regex rxDcw = new(@"DCW Version - (?<v>\S+)", RegexOptions.Compiled);
        private static readonly Regex rxModFw = new(@"Module Firmware Version - (?<v>.+)$", RegexOptions.Compiled);
        private static readonly Regex rxMeterFw = new(@"Meter Firmware Version - (?<v>.+)$", RegexOptions.Compiled);
        private static readonly Regex rxInitDev = new(@"InitializeForDevice majorId (?<maj>[0-9A-F]+)\s+minorId (?<min>[0-9A-F]+)\s+softId (?<soft>[0-9A-F]+)", RegexOptions.Compiled);

        private static readonly string[] PreferredOrder =
            { "ConnectLocalRadio","LocalRadioStatus","Device8052","TunerData","ManufacturingData","AcquireSingleChannelTargetWanSearch" };

        public static List<TechStudioRunRow> ParseRuns(IEnumerable<string> lines)
        {
            string? buildVersion = null;
            var rows = new List<TechStudioRunRow>();
            TechStudioRunRow? cur = null;

            var idToSlot = new Dictionary<string, (int slot, DateTime addedAt)>();
            DateTime? runStart = null; DateTime? runEnd = null;

            void finishRun()
            {
                if (cur == null) return;
                if (runStart.HasValue && runEnd.HasValue)
                    cur.TotalTimeTakenPerRun = (runEnd.Value - runStart.Value).TotalSeconds;
                rows.Add(cur);
                cur = null; idToSlot.Clear(); runStart = runEnd = null;
            }

            int chooseSlotFor(string cmdName, TechStudioRunRow row)
            {
                int byOrder = Array.IndexOf(PreferredOrder, cmdName);
                if (byOrder >= 0 && row.Slots[byOrder].Command == string.Empty) return byOrder;
                for (int i = 0; i < row.Slots.Length; i++)
                    if (row.Slots[i].Command == string.Empty) return i;
                return row.Slots.Length - 1;
            }

            foreach (var line in lines)
            {
                if (buildVersion is null)
                {
                    var mv = rxVersion.Match(line);
                    if (mv.Success) buildVersion = mv.Groups["ver"].Value;
                }

                var ma = rxAdded.Match(line);
                if (ma.Success)
                {
                    var ts = DateTime.ParseExact(ma.Groups["ts"].Value, TS, CultureInfo.InvariantCulture);
                    var cmd = ma.Groups["cmd"].Value;
                    var id = ma.Groups["id"].Value;

                    if (cmd.Equals("ConnectLocalRadio", StringComparison.OrdinalIgnoreCase))
                    {
                        finishRun();
                        cur = new TechStudioRunRow { BuildVersion = buildVersion ?? string.Empty };
                        runStart = ts;
                    }

                    if (cur != null)
                    {
                        int slot = chooseSlotFor(cmd, cur);
                        cur.Slots[slot].Command = cmd;
                        cur.Slots[slot].Timestamp = ts;
                        idToSlot[id] = (slot, ts);
                    }
                    continue;
                }

                var mr = rxMsgIdResp.Match(line);
                if (mr.Success && cur != null)
                {
                    var id = mr.Groups["id"].Value;
                    var ts = DateTime.ParseExact(mr.Groups["ts"].Value, TS, CultureInfo.InvariantCulture);
                    if (idToSlot.TryGetValue(id, out var info))
                    {
                        cur.Slots[info.slot].TimeTakenSeconds = (ts - info.addedAt).TotalSeconds;
                        runEnd = ts;
                    }
                    continue;
                }

                if (cur != null)
                {
                    if (rxSuccessLocal.IsMatch(line) || rxSuccess8052.IsMatch(line) || rxSuccessMfg.IsMatch(line))
                    {
                        for (int i = cur.Slots.Length - 1; i >= 0; i--)
                            if (cur.Slots[i].Command != string.Empty && string.IsNullOrEmpty(cur.Slots[i].Status))
                            { cur.Slots[i].Status = "Success"; break; }
                        continue;
                    }
                    if (rxFailedConnect.IsMatch(line))
                    {
                        cur.Comments = "Failed to connect the endpoint";
                        for (int i = cur.Slots.Length - 1; i >= 0; i--)
                            if (cur.Slots[i].Command != string.Empty && string.IsNullOrEmpty(cur.Slots[i].Status))
                            { cur.Slots[i].Status = "Error"; break; }
                        continue;
                    }

                    var mN = rxNotifyEndpoint.Match(line);
                    if (mN.Success && string.IsNullOrEmpty(cur.LocalRadioId))
                    { cur.LocalRadioId = mN.Groups["id"].Value; continue; }

                    var mE = rxEndpointOk.Match(line);
                    if (mE.Success) { cur.ModuleId = mE.Groups["id"].Value; continue; }

                    if (rxDcw.IsMatch(line)) { cur.DCWVersion = rxDcw.Match(line).Groups["v"].Value; continue; }
                    if (rxModFw.IsMatch(line)) { cur.ModuleFirmware = rxModFw.Match(line).Groups["v"].Value; continue; }
                    if (rxMeterFw.IsMatch(line)) { cur.MeterFirmware = rxMeterFw.Match(line).Groups["v"].Value; continue; }

                    var mDev = rxInitDev.Match(line);
                    if (mDev.Success)
                    {
                        cur.MajorMinorSoftId = $"{mDev.Groups["maj"].Value}/{mDev.Groups["min"].Value}/{mDev.Groups["soft"].Value}";
                        continue;
                    }
                }
            }

            finishRun();
            return rows;
        }

        public static void WriteExcel(List<TechStudioRunRow> runs, string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("MFG API");

            string[] headers = {
                "Build Version","LocalRadio ID","Module ID",
                "Command1","Timestamp","Time Taken (s)","Status",
                "Command2","Timestamp.1","Time Taken (s).1","Status.1",
                "Command3","Timestamp.2","Time Taken (s).2","Status.2",
                "Command4","Timestamp.3","Time Taken (s).3","Status.3",
                "Command5","Timestamp.4","Time Taken (s).4","Status.4",
                "Command6","Timestamp.5","Time Taken (s).5","Status.5",
                "Comments","DCW Version","Module Firmware Version","Meter Firmware Version",
                "Major/Minor/Soft ID","TotalTimeTakenPerRun"
            };

            for (int c = 0; c < headers.Length; c++)
                ws.Cell(1, c + 1).Value = headers[c];

            int row = 2;
            foreach (var r in runs)
            {
                int col = 1;
                ws.Cell(row, col++).Value = r.BuildVersion;
                ws.Cell(row, col++).Value = r.LocalRadioId;
                ws.Cell(row, col++).Value = r.ModuleId;

                for (int i = 0; i < 6; i++)
                {
                    var s = r.Slots[i];
                    ws.Cell(row, col++).Value = s.Command;
                    ws.Cell(row, col++).Value = s.Timestamp;
                    ws.Cell(row, col++).Value = s.TimeTakenSeconds;
                    ws.Cell(row, col++).Value = s.Status;
                }

                ws.Cell(row, col++).Value = r.Comments;
                ws.Cell(row, col++).Value = r.DCWVersion;
                ws.Cell(row, col++).Value = r.ModuleFirmware;
                ws.Cell(row, col++).Value = r.MeterFirmware;
                ws.Cell(row, col++).Value = r.MajorMinorSoftId;
                ws.Cell(row, col++).Value = r.TotalTimeTakenPerRun;

                row++;
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(path);
        }

        public static void ConvertToMfgExcel(string logPath, string excelPath)
        {
            var runs = ParseRuns(File.ReadLines(logPath));
            WriteExcel(runs, excelPath);
        }
    }
}

