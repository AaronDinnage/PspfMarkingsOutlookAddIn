using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace PspfMarkings
{
    [XmlRoot]
    public class Config
    {
        #region Properties

        [XmlElement]
        public string RegexSubject { get; set; }
        [XmlElement]
        public string RegexHeader { get; set; }
        [XmlElement]
        public RegexOptions RegexOptionSet { get; set; }
        [XmlElement]
        public string PspfHeaderName { get; set; }
        [XmlElement]
        public string MsipLabelsHeaderName { get; set; }
        //[XmlElement]
        //public string DefaultOutboundLabel { get; set; }
        //[XmlElement]
        //public string DefaultInboundLabel { get; set; }
        [XmlElement]
        public bool ApplyPspfSubjectToMail { get; set; }
        [XmlElement]
        public bool ApplyPspfXHeaderToMail { get; set; }
        [XmlElement]
        public bool ApplyPspfBodyHeaderToMail { get; set; }
        [XmlElement]
        public bool ApplyPspfSubjectToMeetings { get; set; }
        [XmlElement]
        public bool ApplyPspfXHeaderToMeetings { get; set; }
        [XmlElement]
        public bool ApplyPspfBodyHeaderToMeetings { get; set; }
        [XmlElement]
        public bool RequireMeetingLabel { get; set; }
        [XmlElement]
        public string RequireMeetingLabelMessage { get; set; }
        [XmlArray]
        public ProtectiveMarking[] ProtectiveMarkings { get; set; }

        #endregion

        private static readonly XmlSerializer ConfigurationSerializer = new XmlSerializer(typeof(Config));

        public static Config Current = null;

        public const string Filename = "PspfMarkingsConfig.xml";
        private static readonly string FilePath;

        static Config()
        {
            var assembly = Assembly.GetExecutingAssembly();
            var uriCodeBase = new Uri(assembly.CodeBase);
            string directory = Path.GetDirectoryName(uriCodeBase.LocalPath);

            FilePath = Path.Combine(directory, Filename);

            //CreateDefaultConfig();
            //Save();

            Load();
        }

        public static void CreateDefaultConfig()
        {
            Current = new Config()
            {
                RegexHeader = @"\s*(?:VER=(?<ver>[^,]+))(?:,\s*NS=(?<ns>[^,]+))(?:,\s*SEC=(?<sec>[^,]+))?(?:,\s*DLM=(?<dlm>[^,]+))?(?:,\s*CAVEAT=(?<caveat>[^,]+))*(?:,\s*EXPIRES=(?<expires>[^,]+),\s*DOWNTO=(?<downTo>[^,]+))?(?:,\s*ACCESS=(?<access>[^,]+))*(?:,\s*NOTE=(?<note>[^,]+))?(?:,\s*ORIGIN=(?<origin>[^,]+))?\s*",
                RegexSubject = @"\s*\[(?:SEC=(?<sec>[^,\]]+)|DLM=(?<dlm>[^,\]]+)|SEC=(?<sec>[^,\]]+),\s*DLM=(?<dlm>[^,\]]+))(?:,\s*CAVEAT=(?<caveat>[^,\]]+))*(?:,\s*EXPIRES=(?<expires>[^,\]]+),\s*DOWNTO=(?<downTo>[^,\]]+))?(?:,\s*ACCESS=(?<access>[^,\]]+))*]\s*",
                RegexOptionSet = RegexOptions.CultureInvariant | RegexOptions.IgnoreCase,
                PspfHeaderName = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/x-protective-marking",
                MsipLabelsHeaderName = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/msip_labels/0x0000001F",
                //DefaultInboundLabel = null,
                //DefaultOutboundLabel = null,
                ApplyPspfSubjectToMail = true,
                ApplyPspfXHeaderToMail = true,
                ApplyPspfBodyHeaderToMail = true,
                ApplyPspfSubjectToMeetings = true,
                ApplyPspfXHeaderToMeetings = true,
                ApplyPspfBodyHeaderToMeetings = true,
                RequireMeetingLabel = true,
                RequireMeetingLabelMessage = "This email cannot be sent without a label. Please select:",
                ProtectiveMarkings = new ProtectiveMarking[]
                {
                    new ProtectiveMarking() { DisplayName = "Unofficial", SecurityClassification = "UNOFFICIAL", MsipLabel = "Unofficial", MailBodyHeaderText="UNOFFICIAL", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Official", SecurityClassification = "OFFICIAL", MsipLabel = "Official", MailBodyHeaderText="OFFICIAL", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Official:Sensitive", SecurityClassification = "OFFICIAL:Sensitive", MsipLabel = "Official Sensitive", MailBodyHeaderText="OFFICIAL:Sensitive", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Official:Sensitive (Personal-Privacy)", SecurityClassification = "OFFICIAL:Sensitive", InformationManagementMarkers = new string[] { "Personal-Privacy" }, MsipLabel = "Official Sensitive Personal-Privacy", MailBodyHeaderText="OFFICIAL:Sensitive (Personal-Privacy)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Official:Sensitive (Legal-Privilege)", SecurityClassification = "OFFICIAL:Sensitive", InformationManagementMarkers = new string[] { "Legal-Privilege" }, MsipLabel = "Official Sensitive Legal-Privilege", MailBodyHeaderText="OFFICIAL:Sensitive (Legal-Privilege)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Official:Sensitive (Legislative-Secrecy)", SecurityClassification = "OFFICIAL:Sensitive", InformationManagementMarkers = new string[] { "Legislative-Secrecy" }, MsipLabel = "Official Sensitive Legislative-Secrecy", MailBodyHeaderText="OFFICIAL:Sensitive (Legislative-Secrecy)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Protected", SecurityClassification = "PROTECTED", MsipLabel = "Protected", MailBodyHeaderText="PROTECTED", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Protected (Personal-Privacy)", SecurityClassification = "PROTECTED", InformationManagementMarkers = new string[] { "Personal-Privacy" }, MsipLabel = "Protected Personal-Privacy", MailBodyHeaderText="PROTECTED (Personal-Privacy)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Protected (Legal-Privilege)", SecurityClassification = "PROTECTED", InformationManagementMarkers = new string[] { "Legal-Privilege" }, MsipLabel = "Protected Legal-Privilege", MailBodyHeaderText="PROTECTED (Legal-Privilege)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Protected (Legislative-Secrecy)", SecurityClassification = "PROTECTED", InformationManagementMarkers = new string[] { "Legislative-Secrecy" }, MsipLabel = "Protected Legislative-Secrecy", MailBodyHeaderText="PROTECTED (Legislative-Secrecy)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                    new ProtectiveMarking() { DisplayName = "Protected (Cabinet)", SecurityClassification = "PROTECTED", Caveats = new string[] { "RI:CABINET" }, MsipLabel = "Protected Cabinet", MailBodyHeaderText="PROTECTED (Cabinet)", MailBodyHeaderColour="red", MailBodyHeaderAlign="center", MailBodyHeaderSizePoints="12", MailBodyHeaderFont="helvetica" },
                }
            };
        }

        public static void Load()
        {
            Debug.WriteLine("Config.Load()");

            using (var stream = File.Open(FilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                Current = (Config)ConfigurationSerializer.Deserialize(stream);
        }

        public static void Save()
        {
            Debug.WriteLine("Config.Save()");

            using (var stream = File.Create(FilePath))
                ConfigurationSerializer.Serialize(stream, Current);
        }
    }
}
