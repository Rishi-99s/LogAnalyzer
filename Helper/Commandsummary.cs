using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LogAnalyzer.Helper
{
    public class CommandSummary
    {
        public string CommandName { get; set; }
        public string Timestamp { get; set; }
        public string TimeTaken { get; set; }
        public string Status { get; set; }
    }

    public class DeviceCommandSummary : CommandSummary
    {
        public string BuildVersion { get; set; }
        public string LocalRadioID { get; set; }
        public string ModuleID { get; set; }
        public string DcwVersion { get; set; }
        public string ModuleFirmwareVersion { get; set; }
        public string MeterFirmwareVersion { get; set; }
        public string MajorMinorSoftID { get; set; }
        public string Comments { get; set; }
        public string TotalTimeTakenPerRun { get; set; }
    }
}
