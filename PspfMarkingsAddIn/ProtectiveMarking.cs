using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Xml.Serialization;

namespace PspfMarkings
{
    [DebuggerDisplay("ProtectiveMarking (DisplayName={DisplayName})")]
    [XmlRoot]
    public class ProtectiveMarking
    {
        // Constants
        public const string DefaultVersion = "2018.1";
        public const string PreviousVersion = "2012.3";
        public const string DefaultNamespace = "gov.au";

        [XmlAttribute]
        public string DisplayName { get; set; } = null;

        // Base Attributes
        [XmlAttribute]
        public string Version { get; set; } = DefaultVersion;
        [XmlAttribute]
        public string Namespace { get; set; } = DefaultNamespace;

        // Protective Marking Attributes
        [XmlAttribute]
        public string SecurityClassification { get; set; } = null;
        [XmlElement]
        public string[] Caveats { get; set; } = null;
        [XmlAttribute]
        public string Expires { get; set; } = null;
        [XmlAttribute]
        public string DownTo { get; set; } = null;
        [XmlElement]
        public string[] InformationManagementMarkers { get; set; } = null;
        [XmlAttribute]
        public string Note { get; set; } = null;
        [XmlAttribute]
        public string Origin { get; set; } = null;

        // Deprecated Attributes (2012.3)
        [XmlAttribute]
        public string Dlm { get; set; } = null;

        // MSIP Attributes
        [XmlAttribute]
        public string MsipLabel { get; set; } = null;

        [XmlAttribute]
        public string MailBodyHeaderText { get; set; } = null;
        [XmlAttribute]
        public string MailBodyHeaderColour { get; set; } = null;
        [XmlAttribute]
        public string MailBodyHeaderSizePoints { get; set; } = null;
        [XmlAttribute]
        public string MailBodyHeaderAlign { get; set; } = null;
        [XmlAttribute]
        public string MailBodyHeaderFont { get; set; } = null;


        // Internal Attributes
        [XmlIgnore]
        internal bool IsValid
        {
            get
            {
                if (string.IsNullOrWhiteSpace(SecurityClassification) &&
                    string.IsNullOrWhiteSpace(Dlm))
                    return false;

                return true;
            }
        }

        public ProtectiveMarking() { }

        public string Subject()
        {
            Debug.WriteLine("ProtectiveMarking.Subject()");

            var elements = new List<string>();

            if (!string.IsNullOrWhiteSpace(SecurityClassification))
                elements.Add(string.Format("SEC={0}", SecurityClassification));

            // NOTE: DLM is deprecated
            if (!string.IsNullOrWhiteSpace(Dlm) && Version == PreviousVersion)
                elements.Add(string.Format("DLM={0}", Dlm));

            if (Caveats != null)
                foreach (var caveat in Caveats)
                    elements.Add(string.Format("CAVEAT={0}", caveat));

            if (!string.IsNullOrWhiteSpace(Expires))
            {
                elements.Add(string.Format("EXPIRES={0}", Expires));
                elements.Add(string.Format("DOWNTO={0}", DownTo));
            }

            if (InformationManagementMarkers != null)
                foreach (var access in InformationManagementMarkers)
                    elements.Add(string.Format("ACCESS={0}", access));

            string text = "[" + string.Join(", ", elements) + "]";

            Debug.WriteLine("ProtectiveMarking.Subject() - Result: " + text);

            return text;
        }

        public string Header()
        {
            Debug.WriteLine("ProtectiveMarking.Header()");

            var elements = new List<string>();

            elements.Add(string.Format("VER={0}", Version));
            elements.Add(string.Format("NS={0}", Namespace));
            
            if (!string.IsNullOrWhiteSpace(SecurityClassification))
                elements.Add(string.Format("SEC={0}", SecurityClassification));

            // NOTE: DLM is deprecated
            if (!string.IsNullOrWhiteSpace(Dlm) && Version == PreviousVersion)
                elements.Add(string.Format("DLM={0}", Dlm));

            if (Caveats != null)
                foreach (var caveat in Caveats)
                    elements.Add(string.Format("CAVEAT={0}", caveat));

            if (!string.IsNullOrWhiteSpace(Expires))
            {
                elements.Add(string.Format("EXPIRES={0}", Expires));
                
                if (!string.IsNullOrWhiteSpace(DownTo))
                    elements.Add(string.Format("DOWNTO={0}", DownTo));
            }

            if (InformationManagementMarkers != null)
                foreach (var access in InformationManagementMarkers)
                    elements.Add(string.Format("ACCESS={0}", access));

            if (!string.IsNullOrWhiteSpace(Note))
                elements.Add(string.Format("NOTE={0}", Note));

            if (!string.IsNullOrWhiteSpace(Origin))
                elements.Add(string.Format("ORIGIN={0}", Origin));

            string text = string.Join(", ", elements);

            Debug.WriteLine("ProtectiveMarking.Header() - Result: " + text);

            return text;
        }

        public ProtectiveMarking Clone()
        {
            Debug.WriteLine("ProtectiveMarking.Clone()");

            var protectiveMarking = new ProtectiveMarking()
            {
                DisplayName                     = DisplayName,
                // Current
                Version = Version,
                Namespace                       = Namespace,
                SecurityClassification          = SecurityClassification,
                Caveats                         = Caveats == null ? null : (string[])Caveats.Clone(),
                Expires                         = Expires,
                DownTo                          = DownTo,
                InformationManagementMarkers    = InformationManagementMarkers == null ? null : (string[])InformationManagementMarkers.Clone(),
                Note                            = Note,
                Origin                          = Origin,
                // Deprecated
                Dlm                             = Dlm,
                // Msip
                MsipLabel                       = MsipLabel,
                // Body Header
                MailBodyHeaderAlign             = MailBodyHeaderAlign,
                MailBodyHeaderColour            = MailBodyHeaderColour,
                MailBodyHeaderFont              = MailBodyHeaderFont,
                MailBodyHeaderSizePoints              = MailBodyHeaderSizePoints,
                MailBodyHeaderText              = MailBodyHeaderText,
            };

            return protectiveMarking;
        }

        public static ProtectiveMarking FromRegex(string text, string regex, RegexOptions regexOptions)
        {
            Debug.WriteLine("ProtectiveMarking.FromRegex() - Text=" + text);

            if (string.IsNullOrWhiteSpace(text))
                return null;

            if (string.IsNullOrWhiteSpace(regex))
                return null;

            var match = Regex.Match(text, regex, regexOptions);
            if (!match.Success)
            {
                Debug.WriteLine("ProtectiveMarking.FromRegex() - No match");
                return null;
            }

            if (match.Captures.Count > 1)
                Debug.WriteLine("ProtectiveMarking.FromRegex() - More than one label detected, processing first match");

            return FromRegexMatch(match);
        }

        public static ProtectiveMarking FromRegexMatch(Match match)
        {
            Debug.WriteLine("ProtectiveMarking.FromRegexMatch()");

            if (match == null)
                return null;

            List<string> items;

            var marking = new ProtectiveMarking();

            if (!string.IsNullOrWhiteSpace(match.Groups["ver"].Value))
                marking.Version = match.Groups["ver"].Value;

            if (!string.IsNullOrWhiteSpace(match.Groups["ns"].Value))
                marking.Namespace = match.Groups["ns"].Value;

            if (!string.IsNullOrWhiteSpace(match.Groups["sec"].Value))
                marking.SecurityClassification = match.Groups["sec"].Value;

            var caveats = match.Groups["caveat"].Captures;
            items = new List<string>();
            foreach (Group item in caveats)
                items.Add(item.Value);
            if (items.Count > 0)
                marking.Caveats = items.ToArray();

            if (!string.IsNullOrWhiteSpace(match.Groups["expires"].Value))
                marking.Expires = match.Groups["expires"].Value;

            if (!string.IsNullOrWhiteSpace(match.Groups["downTo"].Value))
                marking.DownTo = match.Groups["downTo"].Value;

            var informationManagementMarkers = match.Groups["access"].Captures;
            items = new List<string>();
            foreach (Group item in informationManagementMarkers)
                items.Add(item.Value);
            if (items.Count > 0)
                marking.InformationManagementMarkers = items.ToArray();

            if (!string.IsNullOrWhiteSpace(match.Groups["note"].Value))
                marking.Note = match.Groups["note"].Value;

            if (!string.IsNullOrWhiteSpace(match.Groups["origin"].Value))
                marking.Origin = match.Groups["origin"].Value;

            // Deprecated (2012.3)
            if (!string.IsNullOrWhiteSpace(match.Groups["dlm"].Value))
            {
                marking.Dlm = match.Groups["dlm"].Value;
                marking.Version = PreviousVersion;
                marking.InformationManagementMarkers = null;
            }

            return marking;
        }

        public static bool Equals(ProtectiveMarking x, ProtectiveMarking y)
        {
            Debug.WriteLine("ProtectiveMarking.Equals()");

            // Compares only the things that matter to equivalence, ignores Notes & Origin (and MSIP attributes) ...

            if (x == null && y == null)
                return true;

            if (x == null && y != null)
                return false;

            if (x != null && y == null)
                return false;

            if (!string.Equals(x.SecurityClassification, y.SecurityClassification, StringComparison.OrdinalIgnoreCase))
                return false;

            if (x.Caveats != null && y.Caveats != null)
            {
                if (x.Caveats.Length != y.Caveats.Length)
                {
                    return false;
                }
                else if (x.Caveats.Length > 0 && y.Caveats.Length > 0)
                {
                    foreach (var caveat in x.Caveats)
                        if (Array.IndexOf(y.Caveats, caveat) == -1)
                            return false;
                }
            }

            if (!string.Equals(x.Expires, y.Expires, StringComparison.OrdinalIgnoreCase))
                return false;

            if (!string.Equals(x.DownTo, y.DownTo, StringComparison.OrdinalIgnoreCase))
                return false;

            if (x.InformationManagementMarkers != null && y.InformationManagementMarkers != null)
            {
                if (x.InformationManagementMarkers.Length != y.InformationManagementMarkers.Length)
                {
                    return false;
                }
                else if (x.InformationManagementMarkers.Length > 0 && y.InformationManagementMarkers.Length > 0)
                {
                    foreach (var marker in x.InformationManagementMarkers)
                        if (Array.IndexOf(y.InformationManagementMarkers, marker) == -1)
                            return false;
                }
            }

            if (!string.Equals(x.Dlm, y.Dlm, StringComparison.OrdinalIgnoreCase))
                return false;

            return true;
        }
    }
}
