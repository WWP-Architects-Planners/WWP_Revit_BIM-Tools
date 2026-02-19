namespace BEPDesigner.WinUI.Models;

public sealed class BepPayload
{
    public string ProjectName { get; set; } = string.Empty;
    public string ProjectNumber { get; set; } = string.Empty;
    public string ProjectAddress { get; set; } = string.Empty;
    public string Client { get; set; } = string.Empty;
    public string ProjectType { get; set; } = string.Empty;
    public string ContractType { get; set; } = string.Empty;
    public string ProjectDescription { get; set; } = string.Empty;
    public string BimLead { get; set; } = string.Empty;
    public string CoordinationMeetingCadence { get; set; } = string.Empty;
    public string PackageMethod { get; set; } = string.Empty;
    public string AutoPublishCadence { get; set; } = string.Empty;
    public string SharingFrequency { get; set; } = string.Empty;
    public string PackageNamingConvention { get; set; } = string.Empty;
    public string GeoCoordinateSystem { get; set; } = string.Empty;
    public string AcquireCoordinatesFromModel { get; set; } = string.Empty;
    public string RevitVersion { get; set; } = string.Empty;
    public string AutoCadVersion { get; set; } = string.Empty;
    public string Civil3DVersion { get; set; } = string.Empty;
    public string DesktopConnectorVersion { get; set; } = string.Empty;
    public string BluebeamVersion { get; set; } = string.Empty;
    public bool EnableWatermark { get; set; }
    public string WatermarkText { get; set; } = "DRAFT";
    public bool LowTrust { get; set; }
    public bool UseConsumedFolder { get; set; }
    public bool RequireRepublish { get; set; }
    public bool StartFresh { get; set; }
    public List<ClashSessionState> Sessions { get; set; } = [];
}
