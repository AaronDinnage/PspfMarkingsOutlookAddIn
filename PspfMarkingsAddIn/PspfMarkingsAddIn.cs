using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace PspfMarkings
{
    public partial class PspfMarkingsAddIn
    {
        #region Constants

        private const string MsipLabelNameStartText = "_Name="; // TODO: Need to collect up multiple name elements?
        private const char MsipSeparator = ';';

        internal const string TemporaryLabelPropertyName = "PspfMarkingsAddIn.Selected";

        #endregion

        #region Static Methods

        private static ProtectiveMarking GetExistingMarking(string pspfHeaderText, string subject)
        {
            ProtectiveMarking existingMarking = null;

            // Preference the header because it contains more information (ie. Origin) ...
            if (!string.IsNullOrWhiteSpace(pspfHeaderText))
                existingMarking = ProtectiveMarking.FromRegex(pspfHeaderText, Config.Current.RegexHeader, Config.Current.RegexOptionSet);

            if (existingMarking == null || !existingMarking.IsValid)
                existingMarking = ProtectiveMarking.FromRegex(subject, Config.Current.RegexSubject, Config.Current.RegexOptionSet);

            return existingMarking;
        }

        private static string GetHeader(Outlook.MeetingItem item, string headerName)
        {
            Outlook.PropertyAccessor accessor = null;

            try
            {
                accessor = item.PropertyAccessor;
                return accessor.GetProperty(headerName);
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                Debug.WriteLine("GetHeader: Failed to retreive header, this is expected behaviour when the header is not present");
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("GetHeader: Unexpected error retrieving header: " + ex.ToString());
            }
            finally
            {
                if (accessor != null)
                    Marshal.ReleaseComObject(accessor);
            }

            return null;
        }

        private static string GetHeader(Outlook.MailItem item, string headerName)
        {
            Outlook.PropertyAccessor accessor = null;

            try
            {
                accessor = item.PropertyAccessor;
                return accessor.GetProperty(headerName);
            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                Debug.WriteLine("GetHeader: Failed to retreive header, this is expected behaviour when the header is not present: " + ex.ToString());
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("GetHeader: Unexpected error retrieving header: " + ex.ToString());
            }
            finally
            {
                if (accessor != null)
                    Marshal.ReleaseComObject(accessor);
            }

            return null;
        }

        private static string GetMsipLabelItem(string msipHeaderText, string labelItem)
        {
            if (string.IsNullOrWhiteSpace(msipHeaderText))
                return null;

            if (string.IsNullOrWhiteSpace(labelItem))
                return null;

            var elements = new List<string>();

            int startIndex;
            int lastPosition = 0;
            while ((startIndex = msipHeaderText.IndexOf(labelItem, lastPosition)) != -1)
            {
                startIndex += labelItem.Length;
                lastPosition = startIndex;

                int endIndex = msipHeaderText.IndexOf(MsipSeparator, startIndex);
                if (endIndex == -1)
                    continue;

                int length = endIndex - startIndex;

                elements.Add(msipHeaderText.Substring(startIndex, length));
            }

            if (elements.Count == 0)
                return null;

            return string.Join(MsipSeparator.ToString(), elements.ToArray());
        }

        private static string ApplyMarkingToSubject(ProtectiveMarking marking, string subject)
        {
            // Remove existing subject marking
            if (!string.IsNullOrWhiteSpace(subject))
                subject = Regex.Replace(subject, Config.Current.RegexSubject, string.Empty, Config.Current.RegexOptionSet);

            // Apply new marking
            if (string.IsNullOrEmpty(subject))
                subject = marking.Subject();
            else
                subject += " " + marking.Subject();

            return subject;
        }

        private static void InsertPspfBodyHeader(Outlook.MailItem item, ProtectiveMarking marking)
        {
            Outlook.Inspector inspector = null;
            dynamic wordEditor = null;

            try
            {
                inspector = item.GetInspector;
                wordEditor = inspector.WordEditor;

                AddMarkingHeaderToDocument(marking, wordEditor);

                // Other methods seen to force the edits to persist:
                //item.Display();
                //inspector.Activate(); // causes the item to pop out a window
            }
            finally
            {
                if (wordEditor != null)
                    Marshal.ReleaseComObject(wordEditor);

                if (inspector != null)
                    Marshal.ReleaseComObject(inspector);
            }
        }

        private static void InsertPspfBodyHeader(Outlook.MeetingItem item, ProtectiveMarking marking)
        {
            Outlook.Inspector inspector = null;
            dynamic wordEditor = null;

            try
            {
                inspector = item.GetInspector;
                wordEditor = inspector.WordEditor;

                AddMarkingHeaderToDocument(marking, wordEditor);

                // Other methods seen to force the edits to persist:
                //item.Display();
                //inspector.Activate(); causes the meeting to pop out a new window and stay open
            }
            finally
            {
                if (wordEditor != null)
                    Marshal.ReleaseComObject(wordEditor);

                if (inspector != null)
                    Marshal.ReleaseComObject(inspector);
            }
        }

        private static void AddMarkingHeaderToDocument(ProtectiveMarking marking, Word.Document document)
        {
            Debug.WriteLine("PspfMarkingsAddIn: AddMarkingHeaderToDocument");
            Debug.WriteLine("==============================================================================");

            Word.Range range = null;
            Word.Font font = null;
            Word.ParagraphFormat format = null;

            try
            {
                object start = 0;
                object end = 0;

                // Move to start
                range = document.Range(ref start, ref end);

                // Insert Paragraph break
                range.InsertParagraphAfter();

                Marshal.ReleaseComObject(range);

                // Move to start
                range = document.Range(ref start, ref end);

                range.Text = marking.MailBodyHeaderText;

                font = range.Font;

                if (!string.IsNullOrWhiteSpace(marking.MailBodyHeaderColour))
                {
                    var colorConverter = new ColorConverter();

                    var color = (Color)colorConverter.ConvertFromString(marking.MailBodyHeaderColour);

                    font.Color = (Word.WdColor)(color.R + 0x100 * color.G + 0x10000 * color.B);
                }

                if (!string.IsNullOrWhiteSpace(marking.MailBodyHeaderSizePoints))
                    font.Size = float.Parse(marking.MailBodyHeaderSizePoints);

                if (!string.IsNullOrWhiteSpace(marking.MailBodyHeaderFont))
                    font.Name = marking.MailBodyHeaderFont;

                if (!string.IsNullOrWhiteSpace(marking.MailBodyHeaderAlign))
                {
                    format = range.ParagraphFormat;

                    switch (marking.MailBodyHeaderAlign.ToLowerInvariant())
                    {
                        case "center":
                        case "centre":
                        case "middle":
                            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                            break;

                        case "left":
                        case "normal":
                            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                            break;

                        case "right":
                            format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
                            break;

                        default:
                            Debug.WriteLine("Unexpected text alignment: " + marking.MailBodyHeaderAlign);
                            break;
                    }
                }

                // Force the edit to persist:
                document.Activate();
            }
            finally
            {
                if (format != null)
                    Marshal.ReleaseComObject(format);

                if (font != null)
                    Marshal.ReleaseComObject(font);

                if (range != null)
                    Marshal.ReleaseComObject(range);
            }
        }

        private static string GetSenderSmtpAddress(Outlook.MailItem item)
        {
            Outlook.Account account = null;

            try
            {
                account = item.SendUsingAccount;
                return account.SmtpAddress;
            }
            finally
            {
                if (account != null)
                    Marshal.ReleaseComObject(account);
            }
        }

        private static string GetSenderSmtpAddress(Outlook.MeetingItem item)
        {
            Outlook.Account account = null;

            try
            {
                account = item.SendUsingAccount;
                return account.SmtpAddress;
            }
            finally
            {
                if (account != null)
                    Marshal.ReleaseComObject(account);
            }
        }

        private static void SetHeader(Outlook.MailItem item, string headerName, string headerValue)
        {
            Outlook.PropertyAccessor accessor = null;

            try
            {
                accessor = item.PropertyAccessor;
                accessor.SetProperty(headerName, headerValue);
            }
            finally
            {
                if (accessor != null)
                    Marshal.ReleaseComObject(accessor);
            }
        }

        private static void SetHeader(Outlook.MeetingItem item, string headerName, string headerValue)
        {
            Outlook.PropertyAccessor accessor = null;

            try
            {
                accessor = item.PropertyAccessor;
                accessor.SetProperty(headerName, headerValue);
            }
            finally
            {
                if (accessor != null)
                    Marshal.ReleaseComObject(accessor);
            }
        }

        private static void UpdateAppointmentSubject(Outlook.MeetingItem item, string subject)
        {
            Outlook.AppointmentItem appointment = null;

            try
            {
                appointment = item.GetAssociatedAppointment(false);
                appointment.Subject = subject;
            }
            finally
            {
                if (appointment != null)
                    Marshal.ReleaseComObject(appointment);
            }
        }

        private static ProtectiveMarking GetUserSelectedMarking(Outlook.MeetingItem item)
        {
            Debug.WriteLine("PspfMarkingsAddIn: GetUserSelectedMarking");
            Debug.WriteLine("==============================================================================");

            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty userProperty = null;

            try
            {
                userProperties = item.UserProperties;
                userProperty = userProperties[TemporaryLabelPropertyName];
                if (userProperty != null)
                    return Config.Current.ProtectiveMarkings[userProperty.Value];
            }
            finally
            {
                if (userProperty != null)
                    Marshal.ReleaseComObject(userProperty);

                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }

            return null;
        }

        private static void DeleteUserSelectedMarking(Outlook.MeetingItem item)
        {
            Debug.WriteLine("PspfMarkingsAddIn: DeleteUserSelectedMarking");
            Debug.WriteLine("==============================================================================");

            Outlook.UserProperties userProperties = null;
            Outlook.UserProperty userProperty = null;

            try
            {
                userProperties = item.UserProperties;
                userProperty = userProperties[TemporaryLabelPropertyName];
                if (userProperty != null)
                    userProperty.Delete();
            }
            finally
            {
                if (userProperty != null)
                    Marshal.ReleaseComObject(userProperty);

                if (userProperties != null)
                    Marshal.ReleaseComObject(userProperties);
            }
        }

        private static ProtectiveMarking PromptUserForLabel()
        {
            var form = new FormLabel();

            var result = form.ShowDialog();

            if (result != DialogResult.OK || form.Selected == null)
            {
                Debug.WriteLine("PromptUserForLabel: User cancelled");
                return null;
            }

            Debug.WriteLine("PromptUserForLabel: New Label Selected");

            var formMarking = Config.Current.ProtectiveMarkings.FirstOrDefault(
                x => string.Equals(x.DisplayName, form.Selected, System.StringComparison.OrdinalIgnoreCase));

            if (formMarking == null)
                throw new System.Exception("PromptUserForLabel: Unexpected error - Failed to locate selected label: " + form.Selected);

            // Clone the marking item so as to not alter the master list of protective markings
            return formMarking.Clone();
        }

        #endregion

        private void PspfMarkingsAddIn_Startup(object sender, System.EventArgs e)
        {
            Debug.WriteLine("PspfMarkingsAddIn: PspfMarkingsAddIn_Startup");
            Debug.WriteLine("==============================================================================");

            if (Config.Current.ProtectiveMarkings.Length == 0)
                throw new System.Exception("Failed to load Protective Markings configuration");

            this.Application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            Debug.WriteLine("PspfMarkingsAddIn: Application_ItemSend");
            Debug.WriteLine("==============================================================================");

            try
            {
                if (item is Outlook.MailItem mailItem)
                {
                    ProcessMailItem(ref cancel, mailItem);
                }
                else if (item is Outlook.MeetingItem meetingItem)
                {
                    ProcessMeetingItem(ref cancel, meetingItem);
                }
                else if (item is Outlook.TaskRequestItem taskRequestItem)
                {
                    ProcessTaskRequestItem(ref cancel, taskRequestItem);
                }
                else
                {
                    Debug.WriteLine("Application_ItemSend: Unrecognized item type");
                }
            }
            catch (System.Exception ex)
            {
                Debug.WriteLine("PspfMarkingsAddIn: Application_ItemSend: " + ex.ToString());
            }
        }

        private void ProcessMailItem(ref bool cancel, Outlook.MailItem item)
        {
            Debug.WriteLine("ProcessMailItem: ProcessMailItem");
            Debug.WriteLine("==============================================================================");

            if (!Config.Current.ApplyPspfBodyHeaderToMail &&
                !Config.Current.ApplyPspfSubjectToMail &&
                !Config.Current.ApplyPspfXHeaderToMail)
            {
                Debug.WriteLine("ProcessMailItem: No Mail labels selected in config.");
                return;
            }

            // TODO: Never works here, need to get it some other way?
            // TODO: trace back to the parent or from the reply creation event?
            //string pspfHeaderText = GetHeader(item, Config.Current.PspfHeaderName);

            string msipHeaderText = GetHeader(item, Config.Current.MsipLabelsHeaderName);
            if (msipHeaderText == null)
            {
                Debug.WriteLine("ProcessMailItem: No MSIP header present");
                return;
            }

            string msipLabelName = GetMsipLabelItem(msipHeaderText, MsipLabelNameStartText);
            if (string.IsNullOrWhiteSpace(msipLabelName))
            {
                Debug.WriteLine("ProcessMailItem: Failed to locate MSIP Label name");
                return;
            }

            var marking = Config.Current.ProtectiveMarkings.FirstOrDefault(
                x => string.Equals(x.MsipLabel, msipLabelName, System.StringComparison.OrdinalIgnoreCase));

            if (marking == null)
            {
                Debug.WriteLine("ProcessMailItem: No Protective Marking aligned to MSIP Label found");
                return;
            }

            // Clone the marking item so as to not alter the master list of protective markings
            marking = marking.Clone();

            // Set the marking origin ...

            //// Only check header, because we only care about the origin
            //var pspfMarking = GetExistingMarking(pspfHeaderText, null);

            //if (ProtectiveMarking.Equals(pspfMarking, marking))
            //    marking.Origin = pspfMarking.Origin;

            if (string.IsNullOrWhiteSpace(marking.Origin))
                marking.Origin = GetSenderSmtpAddress(item);

            if (Config.Current.ApplyPspfSubjectToMail)
                item.Subject = ApplyMarkingToSubject(marking, item.Subject);

            if (Config.Current.ApplyPspfXHeaderToMail)
                SetHeader(item, Config.Current.PspfHeaderName, marking.Header());

            if (Config.Current.ApplyPspfBodyHeaderToMail)
                InsertPspfBodyHeader(item, marking);

            // Save Mail Item to send
            item.Save();
        }

        private void ProcessMeetingItem(ref bool cancel, Outlook.MeetingItem item)
        {
            Debug.WriteLine("ProcessMeetingItem: Outlook.MeetingItem found");
            Debug.WriteLine("==============================================================================");

            if (!Config.Current.ApplyPspfBodyHeaderToMeetings &&
                !Config.Current.ApplyPspfSubjectToMeetings &&
                !Config.Current.ApplyPspfXHeaderToMeetings)
            {
                Debug.WriteLine("ProcessMeetingItem: No Meeting labels selected in config.");
                return;
            }

            ProtectiveMarking marking;

            // TODO: Never works here, need to get it some other way?
            //string pspfHeaderText = GetHeader(item, Config.Current.PspfHeaderName);
            //var pspfMarking = GetExistingMarking(pspfHeaderText, item.Subject);

            var userSelectedMarking = GetUserSelectedMarking(item);
            if (userSelectedMarking == null)
            {
                var pspfMarking = GetExistingMarking(null, item.Subject);

                if (pspfMarking == null || !pspfMarking.IsValid)
                {
                    Debug.WriteLine("ProcessMeetingItem: No Existing Marking");

                    if (!Config.Current.RequireMeetingLabel)
                    {
                        Debug.WriteLine("ProcessMeetingItem: RequireMeetingLabel = false");
                        return;
                    }

                    marking = PromptUserForLabel();
                    if (marking == null)
                    {
                        cancel = true;
                        return;
                    }
                }
                else // pspfMarking is present and valid
                {
                    // even with an existing subject label the xheader and body labels should still be applied.
                    Debug.WriteLine("ProcessMeetingItem: Existing Marking Found");
                    //return; // TODO: Not satisfied with this return outcome.
                    marking = pspfMarking;
                }
            }
            else
            {
                Debug.WriteLine("ProcessMeetingItem: New Label Selected");

                // Clone the marking item so as to not alter the master list of protective markings
                marking = userSelectedMarking.Clone();
            }

            if (string.IsNullOrWhiteSpace(marking.Origin))
                marking.Origin = GetSenderSmtpAddress(item);

            if (Config.Current.ApplyPspfSubjectToMeetings)
            {
                string subject = ApplyMarkingToSubject(marking, item.Subject);
                item.Subject = subject;
                UpdateAppointmentSubject(item, subject);
            }

            if (Config.Current.ApplyPspfXHeaderToMeetings)
                SetHeader(item, Config.Current.PspfHeaderName, marking.Header());

            if (Config.Current.ApplyPspfBodyHeaderToMeetings)
                InsertPspfBodyHeader(item, marking);

            DeleteUserSelectedMarking(item);

            // Save Mail Item to send
            item.Save();
        }

        private void ProcessTaskRequestItem(ref bool cancel, Outlook.TaskRequestItem item)
        {
            Debug.WriteLine("ProcessTaskRequestItem: Outlook.TaskRequestItem found");
            Debug.WriteLine("==============================================================================");

            // TODO: Treat tasks the same as Meetings

        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            Debug.WriteLine("PspfMarkingsAddIn: CreateRibbonExtensibilityObject");
            Debug.WriteLine("==============================================================================");

            return new RibbonLabel();
        }

        private void PspfMarkingsAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(PspfMarkingsAddIn_Startup);
            this.Shutdown += new System.EventHandler(PspfMarkingsAddIn_Shutdown);
        }
        
        #endregion
    }
}
