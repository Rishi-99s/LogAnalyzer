namespace FirstProject.Models
{
    public class TechStudioCommandSlot
    {
        public string Command { get; set; } = string.Empty;
        public DateTime? Timestamp { get; set; }
        public double? TimeTakenSeconds { get; set; }
        public string Status { get; set; } = string.Empty;
    }

    public class TechStudioRunRow
    {
        public string BuildVersion { get; set; } = string.Empty;
        public string LocalRadioId { get; set; } = string.Empty;
        public string ModuleId { get; set; } = string.Empty;

        public TechStudioCommandSlot[] Slots { get; } =
            new TechStudioCommandSlot[] { new(), new(), new(), new(), new(), new() };

        public string Comments { get; set; } = string.Empty;
        public string DCWVersion { get; set; } = string.Empty;
        public string ModuleFirmware { get; set; } = string.Empty;
        public string MeterFirmware { get; set; } = string.Empty;
        public string MajorMinorSoftId { get; set; } = string.Empty;
        public double? TotalTimeTakenPerRun { get; set; }
        public string MajorID { get; set; } = string.Empty;
    }
}
