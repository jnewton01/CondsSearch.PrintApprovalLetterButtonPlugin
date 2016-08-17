using System;

using EllieMae.Encompass.ComponentModel;
using EllieMae.Encompass.Automation;



using System.Net.Mail;
using System.Net;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Linq;
using ConditionsCounters;
using System.Collections;
using EllieMae.EMLite.UI;
using System.Drawing;
using ApprovalsPlugin.Properties;
using EllieMae.Encompass.BusinessObjects.Users;
using EllieMae.Encompass.Collections;
using System.Text;
using Microsoft.Office.Interop.Word;
using System.IO;
using EllieMae.Encompass.BusinessObjects.Loans.Logging;
using System.Text.RegularExpressions;
using FormPlugin;

namespace ConditionsSearch

{
    /// <summary>
    /// Summary description for LoanMonitorPlugin.
    /// </summary>
    [Plugin]
    public class ConditionsCountersPlugin
    {
        // Display the window
        private MonitorWindow currentMonitor = null;

        // The public constructor for the plugin. All plugins must have a public, parameterless
        // constructor. In the constructor you should subscribe to the events you wish to
        // handle within Encompass.

        private Form _mainForm;
        private Form _folderForm;
        private TabControl _tabControl;
        Control[] formControlArray;
        public ConditionsCountersPlugin()

        {

            EncompassApplication.LoanOpened += new EventHandler(EncompassApplication_Login);
        }


        private void EncompassApplication_Login(object sender, EventArgs e)
        {
            try
            {
                foreach (Form openForm in (ReadOnlyCollectionBase)System.Windows.Forms.Application.OpenForms)
                {
                    if (openForm.Text.StartsWith("Encompass360 - ") || openForm.Text.StartsWith("Encompass - "))
                    {
                        this._mainForm = openForm;
                        break;
                    }
                }
                if (this._mainForm == null)
                    throw new Exception("Main Form Not Found");
                Control[] controlArray = this._mainForm.Controls.Find("tabControl", true);
                if (((IEnumerable<Control>)controlArray).Count<Control>() == 0)
                    throw new Exception("Tab Control Not Found");
                this._tabControl = controlArray[0] as TabControl;
                if (this._tabControl == null)
                    throw new Exception("Tab Control is NULL");
                if (this._tabControl.Controls.Count < 2)
                    throw new Exception("Tab Control has too few items");

                Persona UWPersona = EncompassApplication.Session.Users.Personas.GetPersonaByName("Underwriter");
                Persona UWValPersona = EncompassApplication.Session.Users.Personas.GetPersonaByName("Validator");
                Persona AdminPersona = EncompassApplication.Session.Users.Personas.GetPersonaByName("Administrator");
                Persona SuperAdminPersona = EncompassApplication.Session.Users.Personas.GetPersonaByName("Super Administrator");



                if (!EncompassApplication.CurrentUser.Personas.Contains(UWPersona) && !EncompassApplication.CurrentUser.Personas.Contains(UWValPersona) && !EncompassApplication.CurrentUser.Personas.Contains(AdminPersona) && !EncompassApplication.CurrentUser.Personas.Contains(SuperAdminPersona))
                    return;

                this._mainForm.Deactivate += new EventHandler(this._mainFormLostFocus);
            
            }

            catch (Exception ex)
            {
                int num = (int)MessageBox.Show("PipelineHighlighter Cannot Init: " + ex.ToString());
            }
        }

        private void _mainFormLostFocus(object sender, EventArgs e)
        {
            try
            {
                foreach (Form openForm in (ReadOnlyCollectionBase)System.Windows.Forms.Application.OpenForms)
                {
                    if (openForm.Text.Contains("eFolder"))
                    {
                        
                        this._folderForm = openForm;
                        this._folderForm.FormClosing += new FormClosingEventHandler(_folderForm_Closing);
                        break;
                    }
                }

                if (_folderForm == null)
                    return;
                formControlArray = this._folderForm.Controls.Find("pnlMain", true);
           
                if (((IEnumerable<Control>)formControlArray).Count<Control>() == 0)
                    return;
                Panel pnlM = formControlArray[0] as Panel;
              
                    Button btn = new Button();
                btn.Image = Resources.search;
                btn.ImageAlign = ContentAlignment.MiddleLeft;
            
                btn.TextAlign = ContentAlignment.MiddleRight;
                btn.BackColor = Color.Transparent;
                btn.FlatStyle = FlatStyle.Flat;
                btn.FlatAppearance.BorderSize = 0;

                btn.Size = new System.Drawing.Size(135, 21);
                btn.Click += new EventHandler(Btn_Click);

                Button appButtonPrint = new Button();
                appButtonPrint.Image = Resources._1470693069_print_printer;
                appButtonPrint.ImageAlign = ContentAlignment.MiddleLeft;

                appButtonPrint.TextAlign = ContentAlignment.MiddleRight;
                appButtonPrint.BackColor = Color.Transparent;
                appButtonPrint.FlatStyle = FlatStyle.Flat;
                appButtonPrint.FlatAppearance.BorderSize = 0;
               
                appButtonPrint.Size = new System.Drawing.Size(135, 21);
                appButtonPrint.Click += new EventHandler(AppPrintBtn_Click);

                Button TSUMButtonPrint = new Button();
                TSUMButtonPrint.Image = Resources.newprinticon2;
                TSUMButtonPrint.ImageAlign = ContentAlignment.MiddleLeft;

                TSUMButtonPrint.TextAlign = ContentAlignment.MiddleRight;
                TSUMButtonPrint.BackColor = Color.Transparent;
                TSUMButtonPrint.FlatStyle = FlatStyle.Flat;
                TSUMButtonPrint.FlatAppearance.BorderSize = 0;

                TSUMButtonPrint.Size = new System.Drawing.Size(135, 21);
                TSUMButtonPrint.Click += new EventHandler(TSUMPrintBtn_Click);
                GradientPanel p = new GradientPanel();
                FlowLayoutPanel fPnl = new FlowLayoutPanel();
                fPnl.FlowDirection = FlowDirection.RightToLeft;
                p.Name = "newGradPanel";

                fPnl.BackColor = Color.Transparent;
                fPnl.Dock = DockStyle.Fill;
         
                p.Dock = DockStyle.Top;
              
                p.Height = 30;

                if (!pnlM.Controls.ContainsKey("newGradPanel"))
                {
                    fPnl.Controls.Add(btn);
                    Persona UWPersona = EncompassApplication.Session.Users.Personas.GetPersonaByName("Underwriter");
                    if (EncompassApplication.CurrentUser.Personas.Contains(UWPersona))
                    {
                        fPnl.Controls.Add(appButtonPrint);
                        fPnl.Controls.Add(TSUMButtonPrint);
                    }

                       

                    p.Controls.Add(fPnl);
                    pnlM.Controls.Add(p);
                }
           
                
            }
            catch (Exception ex)

            {
                int num = (int)MessageBox.Show("Conditions Search Tool Cannot Init: " + ex.ToString());
            }
        }

        private void TSUMPrintBtn_Click(object sender, EventArgs e)
        {
            File.WriteAllBytes("1008tsum.pdf", Resources.Stripped1008___TSUM_P1);
            File.WriteAllBytes("1008tsum.xml", Resources.Stripped1008___TSUM_P1_pdf);
            GeneratePDFForm PdfForm = new GeneratePDFForm();

            string attachmentFile = PdfForm.GeneratePdfForm(EncompassApplication.CurrentLoan, "1008tsum.pdf", "1008tsum.xml", "_1008___TSUM_P1CLASS");

            EllieMae.Encompass.BusinessObjects.Loans.Attachment att = EncompassApplication.CurrentLoan.Attachments.AddImage(attachmentFile);

            LogEntryList lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("1008 Completed By U/W");
            if (lst.Count == 0)
            {
                EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add("1008 Completed By U/W", "Approval");
                att.Title = "1008 Transmital Summary " + DateTime.Now.ToShortDateString();
                lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("1008 Completed By U/W");
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                appLetter.Attach(att);
            }

            if (lst.Count > 0)
            {
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                att.Title = "1008 Transmital Summary " + DateTime.Now.ToShortDateString();
                appLetter.Attach(att);
            }

        
    }


        private void _folderForm_Closing(object sender, EventArgs e)
        {
            if (currentMonitor != null)
            {
                currentMonitor.Close();
            }

            }           

        private void Btn_Click(object sender, EventArgs e)
        {
      
            currentMonitor = new MonitorWindow();
            currentMonitor.Show();
     
        }
        //get bucket by name and create the bucket if not in loan already.

        private TrackedDocument getBucket(string bucketTitle)
        {
            LogEntryList findings = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle(bucketTitle);
            if (findings.Count < 1)
            {
                EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add(bucketTitle, "Submittal");
                findings = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle(bucketTitle);
            }

            return (TrackedDocument)findings[0];
        }

        private void removeOldApprovalLetters()
        {
            TrackedDocument junkDocs = getBucket("* Junk Folder");
            //go through all the buckets and get the approval letters bucket and get all attachments then remove them
            foreach (TrackedDocument doc in EncompassApplication.CurrentLoan.Log.TrackedDocuments)
            {
                if (doc.Title.Equals("* Approval Letter"))
                {
                    TrackedDocument approvalDocs = doc;
                    foreach (EllieMae.Encompass.BusinessObjects.Loans.Attachment appAtt in approvalDocs.GetAttachments())
                    {
                        doc.Detach(appAtt);
                        junkDocs.Attach(appAtt);
                    }

                }
            }
        }

        private void suspenseLetter()
        {
            string fileName = "Suspense Letter.pdf";
            string attachmentFile = ToPdf("_suspendedLetter.docx");

            EllieMae.Encompass.BusinessObjects.Loans.Attachment att = EncompassApplication.CurrentLoan.Attachments.AddImage(attachmentFile);

         
            LogEntryList lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");

            if (lst.Count == 0)
            {
                EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add("* Approval Letter", "Cond'l Approval");
                att.Title = "Suspense Letter  " + DateTime.Now.ToShortDateString();
                lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                appLetter.Attach(att);
            }

            if (lst.Count > 0)
            {
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                att.Title = "Suspense Letter  " + DateTime.Now.ToShortDateString();

                appLetter.Attach(att);
            }
           
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanProcessorID).Email.ToString(), "Loan Suspended - " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString(), attachmentFile, fileName, suspenedAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanOfficerID).Email.ToString(), "Loan Suspended - " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString(), attachmentFile, fileName, suspenedAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.Session.UserID).Email.ToString(), "Loan Suspended - " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString(), attachmentFile, fileName, suspenedAppBody());
            SendMail("jnewton@gsfmail.com", "Loan Suspended - " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString(), attachmentFile, fileName, suspenedAppBody());

        }
        private void ctcLetter()
        {
            string fileName = "Clear to Close.pdf";
            string attachmentFile = ToPdf("_ctcLetter.docx");

            EllieMae.Encompass.BusinessObjects.Loans.Attachment att = EncompassApplication.CurrentLoan.Attachments.AddImage(attachmentFile);
          
            LogEntryList lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");
            if (lst.Count == 0)
            {
                EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add("* Approval Letter", "Approval");
                att.Title = "Clear to Close Letter " + DateTime.Now.ToShortDateString();
                lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                appLetter.Attach(att);
            }

            if (lst.Count > 0)
            {
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                att.Title = "Clear to Close Letter " + DateTime.Now.ToShortDateString();
                appLetter.Attach(att);
            }
           
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanProcessorID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " is clear to close!", attachmentFile, fileName, ctcAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanOfficerID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " is clear to close!", attachmentFile, fileName, ctcAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.Session.UserID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " is clear to close!", attachmentFile, fileName, ctcAppBody());
            SendMail("jnewton@gsfmail.com", "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " is clear to close!", attachmentFile, fileName, ctcAppBody());
            
        }
        private void cdlLetter()
        {
            string fileName = "Cond'l Approval Letter.pdf";
            string attachmentFile = ToPdf("_capprovalLetter.docx");
           
            EllieMae.Encompass.BusinessObjects.Loans.Attachment att = EncompassApplication.CurrentLoan.Attachments.AddImage(attachmentFile);
            
            LogEntryList lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");
            if (lst.Count == 0)
            {
                EncompassApplication.CurrentLoan.Log.TrackedDocuments.Add("* Approval Letter", "Approval");
                att.Title = "Cond'l Approval Letter " + DateTime.Now.ToShortDateString();
                lst = EncompassApplication.CurrentLoan.Log.TrackedDocuments.GetDocumentsByTitle("* Approval Letter");
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                appLetter.Attach(att);
            }

            if (lst.Count > 0)
            {
                TrackedDocument appLetter = (TrackedDocument)lst[0];
                att.Title = "Cond'l Approval Letter " + DateTime.Now.ToShortDateString();
                appLetter.Attach(att);
            }
           
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanProcessorID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " has been conditionally approved!", attachmentFile, fileName, cdlAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.LoanOfficerID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " has been conditionally approved!", attachmentFile, fileName, cdlAppBody());
            SendMail(EncompassApplication.Session.Users.GetUser(EncompassApplication.CurrentLoan.Session.UserID).Email.ToString(), "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " has been conditionally approved!", attachmentFile, fileName, cdlAppBody());
            SendMail("jnewton@gsfmail.com", "Congratulations! The loan for " + EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString() + " " + EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString() + " has been conditionally approved!", attachmentFile, fileName, cdlAppBody());

        }

        private void AppPrintBtn_Click(object sender, EventArgs e)
        {
          
            removeOldApprovalLetters();
           

            if (EncompassApplication.CurrentLoan.Fields["CX.UW.SUSPEND"].Value.ToString().Equals("X"))
            {
                suspenseLetter();
            }
            if (EncompassApplication.CurrentLoan.Fields["CX.UWI.CTC"].Value.ToString().Equals("X"))
            {
                ctcLetter();
            }

            if (EncompassApplication.CurrentLoan.Fields["CX.UWI.CTC"].Value.ToString().Equals("") && EncompassApplication.CurrentLoan.Fields["CX.UW.SUSPEND"].Value.ToString().Equals(""))
            {
                cdlLetter();
            }

                MessageBox.Show("Your Decision Letter has been placed in the eFolder");


        }

        public string cdlAppBody()
        {

            StringBuilder str = new StringBuilder();
            str.Clear();
            str.Append(Resources.cdlAppBody);
            try
            {
                str.Replace("<<M_Loan_Pro>>", EncompassApplication.CurrentLoan.Fields["362"].FormattedValue.ToString());
                str.Replace("<<M_37>>", EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString());
                str.Replace("<<M_36>>", EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString());
                str.Replace("<<Loan_Number_364>>", EncompassApplication.CurrentLoan.Fields["364"].FormattedValue.ToString());
                str.Replace("<<Subject_Property_Address_11>>", EncompassApplication.CurrentLoan.Fields["11"].FormattedValue.ToString());
                str.Replace("<<Note_Rate_3>>", EncompassApplication.CurrentLoan.Fields["3"].FormattedValue.ToString());
                str.Replace("<<M_984>>", EncompassApplication.CurrentLoan.Fields["984"].FormattedValue.ToString());
                str.Replace("<<M_CX.UINOTES>>", underwriterNotes(EncompassApplication.CurrentLoan.Fields["CX.UWI.UNDERWRITERNOTES"].FormattedValue.ToString()));

            }

            catch (Exception ex)
            {

            }


            return str.ToString();

        }

        public string ctcAppBody()
        {
            StringBuilder str = new StringBuilder();
            str.Clear();
            str.Append(Resources.ctcAppBody);
            try
            {

                str.Replace("<<M_Loan_Pro>>", EncompassApplication.CurrentLoan.Fields["362"].FormattedValue.ToString());
                str.Replace("<<M_37>>", EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString());
                str.Replace("<<M_36>>", EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString());
                str.Replace("<<Loan_Number_364>>", EncompassApplication.CurrentLoan.Fields["364"].FormattedValue.ToString());
                str.Replace("<<Subject_Property_Address_11>>", EncompassApplication.CurrentLoan.Fields["11"].FormattedValue.ToString());
                str.Replace("<<Note_Rate_3>>", EncompassApplication.CurrentLoan.Fields["3"].FormattedValue.ToString());
                str.Replace("<<M_984>>", EncompassApplication.CurrentLoan.Fields["984"].FormattedValue.ToString());
                str.Replace("<<cx.uw.mustclosedate>>", EncompassApplication.CurrentLoan.Fields["CX.UW.CLOSEBYDATE"].FormattedValue.ToString());
                str.Replace("<<M_CX.UINOTES>>", underwriterNotes(EncompassApplication.CurrentLoan.Fields["CX.UWI.UNDERWRITERNOTES"].FormattedValue.ToString()));
            }
            catch (Exception ex)
            {

            }

            return str.ToString();
        }
        public string suspenedAppBody()
        {
            StringBuilder str = new StringBuilder();
            str.Clear();
            str.Append(Resources.suspendedAppBody);
            try
            {

                str.Replace("<<M_Loan_Pro>>", EncompassApplication.CurrentLoan.Fields["362"].FormattedValue.ToString());
                str.Replace("<<M_37>>", EncompassApplication.CurrentLoan.Fields["4000"].FormattedValue.ToString());
                str.Replace("<<M_36>>", EncompassApplication.CurrentLoan.Fields["4002"].FormattedValue.ToString());
                str.Replace("<<Loan_Number_364>>", EncompassApplication.CurrentLoan.Fields["364"].FormattedValue.ToString());
                str.Replace("<<Subject_Property_Address_11>>", EncompassApplication.CurrentLoan.Fields["11"].FormattedValue.ToString());
                str.Replace("<<Note_Rate_3>>", EncompassApplication.CurrentLoan.Fields["3"].FormattedValue.ToString());
                str.Replace("<<M_984>>", EncompassApplication.CurrentLoan.Fields["984"].FormattedValue.ToString());
                str.Replace("<<M_CX.UINOTES>>", underwriterNotes(EncompassApplication.CurrentLoan.Fields["CX.UWI.UNDERWRITERNOTES"].FormattedValue.ToString()));
            }

            catch (Exception ex)
            {

            }

            return str.ToString();
        }

        /// <summary>
        /// Creates the clear to close letter body html email message.
        /// </summary>
        /// <returns>System.String.</returns>
 
        private string underwriterNotes(string input)
        {

            Regex rgx = new Regex("([\n])+");
            string result = rgx.Replace(input, "<br>");
            return result;
        }

        private void SendMail(string EmailAddress, string Subject, string attachmentName,string fileName)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage("gsencompass@gsfmail.com", EmailAddress);

            //  The src attribute for the image tag is set to the filePathInHtml:
            System.Net.Mail.Attachment attachment;
           
            attachment = new System.Net.Mail.Attachment(attachmentName);
            mail.IsBodyHtml = false;
            mail.Body = "";

            attachment.Name = fileName;
            mail.Attachments.Add(attachment);
            SmtpClient client = new SmtpClient();

            client.Credentials = new NetworkCredential("gsencompass@gsfmail.com", "Sup3rSp33d1$");
            client.EnableSsl = true;
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Host = "smtp.office365.com";

            mail.Subject = Subject;

            client.Send(mail);
        }

        private string GetUWCenter()
        {
            string UWCenter = "Ann Arbor";
            switch (EncompassApplication.CurrentUser.Email)
            {
                case "jenetanya@gsfmail.com":
                    UWCenter = "Glendale";
                    break;
                case "jnicolas@gsfmail.com":
                    UWCenter = "Glendale";
                    break;
                case "lvillero@gsfmail.com":
                    UWCenter = "Glendale";
                    break;
                case "slakovic@gsfmail.com":
                    UWCenter = "Glendale";
                    break;
            }

            return UWCenter;
            // Organization orgs = EncompassApplication.Session.Organizations.GetOrganization(EncompassApplication.CurrentUser.OrganizationID)
            //   return EncompassApplication.CurrentUser.OrganizationID.ToString();
        }

        public string ToPdf(string fileName)
        {
            File.WriteAllBytes("_capprovalLetter.docx", Resources.__capprovalLetter);
            File.WriteAllBytes("_suspendedLetter.docx", Resources._suspendedLetter);
            File.WriteAllBytes("_ctcLetter.docx", Resources._ctcLetter);

         
            string filePath;
            filePath = Environment.GetEnvironmentVariable("temp").ToString() + "\\" + Path.GetRandomFileName() + EncompassApplication.CurrentLoan.Guid.ToString() + ".pdf";

            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document wordDoc = new Microsoft.Office.Interop.Word.Document();
            Object oMissing = System.Reflection.Missing.Value;

            word.Visible = false;

            Object filepath = @"c:\SmartClientCache\Apps\Ellie Mae\Encompass\" + fileName;
            Object confirmconversion = System.Reflection.Missing.Value;
            Object readOnly = false;
            Object obOpenAndRepair = false;
            Object saveto = filePath;
            Object oallowsubstitution = System.Reflection.Missing.Value;


            wordDoc = word.Documents.Open(ref filepath, ref confirmconversion, ref readOnly, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                                          ref oMissing, ref oMissing, ref oMissing, ref oMissing);

            wordDoc.Activate();
            Dictionary<string, string> dic = new Dictionary<string, string>()

    {
                { "«M_cxdotlosdotinvestor»", "CX.LOS.INVESTOR"},
                { "«M_cxdotuwdotiddotcreditrefreshexp»", "CX.UW.ID.CREDITREFRESHEXP"},
                { "«VF_Log_MS_Date_Post__20Closing»", "Log.MS.Date.Post Closing"},
                { "«M_cxdotuwdotcondapp1»", "CX.UW.CONDAPP1"},
                { "«M_cxdotuwdotpaexpires»", "CX.UW.PAEXPIRES"},
                { "«M_cxdotuwdotiddotcincomeexp»", "CX.UW.ID.CINCOMEEXP"} ,
                { "«M_cxdotuwdotiddotcplexp»", "CX.UW.ID.CPLEXP"},
                { "«M_cxdotuwdotiddotassetsexp»", "CX.UW.ID.ASSESTSEXP"},
                { "«M_cxdotuwdotiddotcreditexp»", "CX.UW.ID.CREDITEXP"},
                { "«M_cxdotuwdotiddotbincomeexp»", "CX.UW.ID.BINCOMEEXP"},
                { "«M_cxdotuwdotiddotcovoe1exp»", "CX.UW.ID.CVOE1EXP"},
                { "«M_cxdotuwdotiddotappraisalexp»", "CX.UW.ID.APPRAISALEXP"},
                { "«M_cxdotuwdotiddotvoe1exp»", "CX.UW.ID.VOE1EXP"},

                { "«M_cxdotuwdotclosebydate»", "CX.UW.CLOSEBYDATE"},

                { "«M_1256»", "1256"},
                { "«M_1262»", "1262"},
                { "«M_1263»", "1263"},
                { "«Loan_Number_364»", "364"},
                { "«M_1014»", "1014"},
                { "«M_976»", "976"},
                { "«M_984»", "984"},
                { "«Note_Rate_3»", "3"},
                { "«M_2151»", "2151"},
                { "«M_2»", "2"},
                { "«M_1401»", "1401"},
                { "«M_140»", "140"},
                { "«M_19»", "19"},
                { "«M_3»", "3"},
                { "«M_1172»", "1172"},
                { "«M_136»", "136"} ,
                { "«M_1811»", "1811"} ,
                { "«M_356»", "356"} ,
                { "«M_608»", "608"},
                { "«M_353»", "353"},
                { "«M_912»", "912"} ,
                { "«M_2217»", "2217"},
                { "«M_1733»", "1733"},
                { "«M_2293»", "2293"} ,
                { "«M_1742»", "1742"} ,
                { "«M_2294»", "2294"} ,

                { "«M_978»", "978"},
                { "«M_4»", "4"},
                { "«M_325»", "325"} ,
                { "«M_420»", "420"} ,
                { "«M_689»", "689"} ,
                { "«M_740»", "740"},
                { "«M_742»", "742"},
                { "«M_MORNETdotX67»", "MORNET.X67"},
                { "«M_1389»", "1389"},
                { "«M_2216»", "2216"},
                { "«M_SERVICEdotX7»", "SERVICE.X7"},
                { "«M_SERVICEdotX1»", "SERVICE.X1"},
                { "«M_SERVICEdotX13»", "SERVICE.X13"},
                { "«M_SERVICEdotX42»", "SERVICE.X42"} ,
                { "«M_SERVICEdotX44»", "SERVICE.X44"} ,
                { "«M_SERVICEdotX57»", "SERVICE.X57"},
                { "«M_SERVICEdotX81»", "SERVICE.X81"} ,
                { "«M_SERVICEdotX17»", "SERVICE.X17"} ,
                { "«M_SERVICEdotX14»", "SERVICE.X14"},
                { "«M_SERVICEdotX82»", "SERVICE.X82"},
                { "«M_SERVICEdotX20»", "SERVICE.X20"} ,
                { "«M_SERVICEdotX24»", "SERVICE.X24"},
                { "«M_VENDdotX178»", "VEND.X178"},
                { "«M_VENDdotX179»", "VEND.X179"},
                { "«M_VENDdotX180»", "VEND.X180"} ,
                { "«M_VENDdotX181»", "VEND.X181"},
                { "«M_VENDdotX182»", "VEND.X182"},
                { "«M_13»", "13"},
                { "«M_682»", "682"},

                { "«M_SERVICEdotX2»", "SERVICE.X2"},
                { "«M_SERVICEdotX3»", "SERVICE.X3"},
                { "«M_SERVICEdotX4»", "SERVICE.X4"},
                { "«M_SERVICEdotX5»", "SERVICE.X5"},
                { "«M_SERVICEdotX6»", "SERVICE.X6"},


                { "«VF_Log_MS_Date_Cond__27l__20Approval»","Log.MS.Date.Cond'l Approval" },

                { "«M_SERVICEdotX32»", "SERVICE.X32"},
                { "«M_SERVICEdotX34»", "SERVICE.X34"},
                { "«M_SERVICEdotX35»", "SERVICE.X35"},
                { "«M_SERVICEdotX36»", "SERVICE.X36"},
                { "«M_SERVICEdotX37»", "SERVICE.X37"},
                { "«M_SERVICEdotX38»", "SERVICE.X38"},
                { "«M_SERVICEdotX33»", "SERVICE.X33"},
                { "«M_SERVICEdotX15»", "SERVICE.X15"},
                { "«M_SERVICEdotX26»", "SERVICE.X26"},
                { "«M_SERVICEdotX25»", "SERVICE.X25"},


                { "«Subject_Property_Address_11»", "11"},
                { "«Subject_Property_County_13»", "13"},
                { "«M_12»", "12"},
                { "«M_14»", "14"},
                { "«M_15»", "15"},
                { "«M_1553»", "1041"},
                { "«Loan_Purpose_19»", "19"},
                { "«M_CUST20FV»", "CUST20FV"},


                { "«M_1414»", "1414"},
                { "«M_37»", "4000"},
                { "«M_36»", "4002"},
                { "«M_67»", "67"},
                { "«M_1450»", "1450"},
                { "«M_69»", "4004"},
                { "«M_68»", "4006"},
                { "«M_60»", "60"},
                { "«M_1452»", "1452"},
                { "«M_1415»", "1415"},
                { "«M_11»", "11"},
                { "«M_4000»", "4000"},
                { "«M_4002»", "4002"},
                { "«M_FR0104»", "FR0104"},
                { "«M_FR0106»", "FR0106"},
                { "«M_FR0107»", "FR0107"},
                { "«M_FR0108»", "FR0108"},



    };




            foreach (Microsoft.Office.Interop.Word.Field tmpRange in wordDoc.Fields)
            {


                switch (tmpRange.Result.Text)
                {
                    case "«VF_UWC_PTA»":
                        tmpRange.Result.Text = PTAConditions();
                        break;
                    case "«VF_UWC_PTD»":
                        tmpRange.Result.Text = PTDConditions();
                        break;
                    case "«VF_UWC_AC»":
                        tmpRange.Result.Text = AtClosingConditions();
                        break;
                    case "«VF_UWC_PTF»":
                        tmpRange.Result.Text = PTFConditions();
                        break;
                    case "«M_UW_CENTER»":
                        tmpRange.Result.Text = GetUWCenter();
                        break;
                    default:

                        if (tmpRange.Result.Text != null)
                        {
                            if (dic.ContainsKey(tmpRange.Result.Text) == true)
                            {

                                string dicResults = dic[tmpRange.Result.Text];
                                try
                                {
                                    tmpRange.Result.Text = EncompassApplication.CurrentLoan.Fields[dicResults].FormattedValue.ToString();

                                }

                                catch (Exception ex)
                                {
                                    tmpRange.Result.Text = tmpRange.Result.Text;
                                }
                            }
                        }
                        break;

                }





            }
            object fileFormat = WdSaveFormat.wdFormatPDF;

            try
            {
                wordDoc.SaveAs(ref saveto, ref fileFormat, ref oMissing, ref oMissing, ref oMissing,
                               ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                               ref oMissing, ref oMissing, ref oMissing, ref oallowsubstitution, ref oMissing,
                               ref oMissing);

                wordDoc.Close();
                word.Quit();

            }
            catch (Exception ex)
            {
                wordDoc.Close();
                word.Quit();
            }


          
            File.Delete("_capprovalLetter.docx");
            File.Delete("_suspendedLetter.docx");
            File.Delete("_ctcLetter.docx");


            return filePath;

        }


        /// <summary>
        /// Returns the conditions that equal external use only and have a status that requires to be completed still.
        /// </summary>
        /// <returns>System.String.</returns>
        private string PTAConditions()
        {
            StringBuilder conditions = new StringBuilder();

            foreach (EllieMae.Encompass.BusinessObjects.Loans.Logging.UnderwritingCondition cond in EncompassApplication.CurrentLoan.Log.UnderwritingConditions)
            {
                if (cond.PriorTo.Equals("PTA") && cond.ForExternalUse)
                {
                    if (cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Added) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Expected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.PastDue) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rejected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rerequested) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Requested))
                    {
                        conditions.AppendLine(cond.Title.ToString() + cond.Description.ToString());
                    }
                }
            }
            return conditions.ToString();
        }

        /// <summary>
        /// Returns the conditions that equal external use only and have a status that requires to be completed still.
        /// </summary>
        /// <returns>System.String.</returns>
        private string PTDConditions()
        {
            StringBuilder conditions = new StringBuilder();

            foreach (EllieMae.Encompass.BusinessObjects.Loans.Logging.UnderwritingCondition cond in EncompassApplication.CurrentLoan.Log.UnderwritingConditions)
            {
                if (cond.PriorTo.Equals("PTD") && cond.ForExternalUse)
                {
                    if (cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Added) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Expected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.PastDue) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rejected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rerequested) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Requested))
                    {

                        conditions.AppendLine(cond.Title.ToString() + cond.Description.ToString());
                    }
                }

            }
            return conditions.ToString();
        }

        /// <summary>
        /// Returns the conditions that equal external use only and have a status that requires to be completed still.
        /// </summary>
        /// <returns>System.String.</returns>
        private string PTFConditions()
        {
            StringBuilder conditions = new StringBuilder();

            foreach (EllieMae.Encompass.BusinessObjects.Loans.Logging.UnderwritingCondition cond in EncompassApplication.CurrentLoan.Log.UnderwritingConditions)
            {
                if (cond.PriorTo.Equals("PTF"))
                {
                    if (cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Added) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Expected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.PastDue) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rejected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rerequested) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Requested))
                    {
                        conditions.AppendLine(cond.Title.ToString() + cond.Description.ToString());
                    }
                }

            }
            return conditions.ToString();
        }

        /// <summary>
        /// Returns the conditions that equal external use only and have a status that requires to be completed still.
        /// </summary>
        /// <returns>System.String.</returns>
        private string AtClosingConditions()
        {
            StringBuilder conditions = new StringBuilder();

            foreach (EllieMae.Encompass.BusinessObjects.Loans.Logging.UnderwritingCondition cond in EncompassApplication.CurrentLoan.Log.UnderwritingConditions)
            {
                if (cond.PriorTo.Equals("AC"))
                {
                    if (cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Added) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Expected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.PastDue) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rejected) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Rerequested) || cond.Status.Equals(EllieMae.Encompass.BusinessObjects.Loans.Logging.ConditionStatus.Requested))
                    {
                        conditions.AppendLine(cond.Title.ToString() + cond.Description.ToString());
                    }
                }
            }
            return conditions.ToString();
        }


        /// <summary>
        /// Sends the mail.
        /// </summary>
        /// <param name="EmailAddress">The email address.</param>
        /// <param name="Subject">The subject.</param>
        /// <param name="attachmentName">Name of the attachment.</param>
        /// <param name="fileName">Name of the file.</param>
        /// <param name="bodyFileName">Name of the body file.</param>
        private void SendMail(string EmailAddress, string Subject, string attachmentName, string fileName, string bodyFileName)
        {
            System.Net.Mail.MailMessage mail = new System.Net.Mail.MailMessage("gsencompass@gsfmail.com", EmailAddress);

            //  The src attribute for the image tag is set to the filePathInHtml:
            System.Net.Mail.Attachment attachment;

            attachment = new System.Net.Mail.Attachment(attachmentName);
            mail.IsBodyHtml = true;
            mail.Body = bodyFileName;

            attachment.Name = fileName;
            mail.Attachments.Add(attachment);
            SmtpClient client = new SmtpClient();

            client.Credentials = new NetworkCredential("gsencompass@gsfmail.com", "Sup3rSp33d1$");
            client.EnableSsl = true;
            client.Port = 587;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.Host = "smtp.office365.com";

            mail.Subject = Subject;

            client.Send(mail);
        }

    }
}
