using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace PspfMarkings
{
    [ComVisible(true)]
    public class RibbonLabel : Office.IRibbonExtensibility
    {
        public const string RibbonLabelXmlFile = "RibbonLabel.xml";
        public const string ButtonIdPrefix = "button";

        public RibbonLabel()
        {
            Debug.WriteLine("RibbonLabel()");
            Debug.WriteLine("==============================================================================");
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            Debug.WriteLine("RibbonLabel: GetCustomUI");
            Debug.WriteLine("==============================================================================");

            if (ribbonID == "Microsoft.Outlook.Appointment")
            {
                Debug.WriteLine("RibbonLabel: GetCustomUI - Microsoft.Outlook.Appointment");
                Debug.WriteLine("==============================================================================");

                var assembly = Assembly.GetExecutingAssembly();
                var uriCodeBase = new System.Uri(assembly.CodeBase);
                string directory = Path.GetDirectoryName(uriCodeBase.LocalPath);
                string filePath = Path.Combine(directory, RibbonLabelXmlFile);

                return File.ReadAllText(filePath);
            }

            return string.Empty;
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            Debug.WriteLine("RibbonLabel: Ribbon_Load");
            Debug.WriteLine("==============================================================================");
        }

        public string MenuLabel_GetContent(Office.IRibbonControl control)
        {
            Debug.WriteLine("RibbonLabel: MenuLabel_GetContent");
            Debug.WriteLine("==============================================================================");

            var menu = new StringBuilder();
            menu.AppendLine(@"<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">");

            int index = 0;
            foreach (var marking in Config.Current.ProtectiveMarkings)
                menu.AppendLine(string.Format(@"<button id=""{0}{1}"" label=""{2}"" onAction=""MenuLabel_ButtonAction"" />", ButtonIdPrefix, index++, marking.DisplayName));

            menu.AppendLine(@"</menu>");
            return menu.ToString();
        }

        public void MenuLabel_ButtonAction(Office.IRibbonControl control)
        {
            Debug.WriteLine("RibbonLabel: MenuLabel_ButtonAction");
            Debug.WriteLine("==============================================================================");

            string buttonIndexText = control.Id.Substring(ButtonIdPrefix.Length);
            int selectedItemIndex = int.Parse(buttonIndexText);

            object context = null;
            object currentItem = null;
            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty userProperty = null;

            try
            {
                context = control.Context;
                var inspector = (Outlook.Inspector)context;
                currentItem = inspector.CurrentItem;

                if (currentItem is Outlook.AppointmentItem item)
                {
                    // Set selected item index to temporary label property to be read on send
                    userProperties = item.UserProperties;
                    userProperty = userProperties.Add(PspfMarkingsAddIn.TemporaryLabelPropertyName, Outlook.OlUserPropertyType.olInteger);
                    userProperty.Value = selectedItemIndex;
                    //item.Save();

                    UpdateSubject(item, selectedItemIndex);
                }
            }
            finally
            {
                if (userProperty != null)
                    Marshal.ReleaseComObject(userProperty);

                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);

                if (currentItem != null)
                    Marshal.ReleaseComObject(currentItem);

                if (context != null)
                    Marshal.ReleaseComObject(context);
            }
        }

        private static void UpdateSubject(Outlook.AppointmentItem item, int selectedItemIndex)
        {
            Debug.WriteLine("RibbonLabel: UpdateSubject");
            Debug.WriteLine("==============================================================================");

            var marking = Config.Current.ProtectiveMarkings[selectedItemIndex];

            // TODO: Force the subject field to save before any changes are mode - Unintended consequence is the draft meeting is saved to the calendar.
            //item.Save();

            // Remove existing subject marking
            if (!string.IsNullOrEmpty(item.Subject))
                item.Subject = Regex.Replace(item.Subject, Config.Current.RegexSubject, string.Empty, Config.Current.RegexOptionSet);

            // Apply new subject marking
            if (string.IsNullOrEmpty(item.Subject))
                item.Subject = marking.Subject();
            else
                item.Subject += " " + marking.Subject();
        }

        #endregion
    }
}
