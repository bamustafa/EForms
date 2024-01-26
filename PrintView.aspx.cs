using Microsoft.SharePoint;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Winnovative.WnvHtmlConvert;

public partial class PrintView : System.Web.UI.Page
{
    #region GENERIC EMAIL TEMPLATE VARIABLES

    string CurrentUserEmail = "",
           Title = "",
           EmailType = "",
           EmailValue = "",
           ListName = "",

           HeaderChunk = "",
           EditedChunk = "",
           FooterChunck = "",
           AddNewRow = "",
           ColumnInBrackets = "",
           value = "";

    bool isMulti = false,
         isNewRow = false,
         isFirstVisit = false;

    #endregion

    protected void Page_PreInit(object sender, EventArgs e)
    {
        string MasterPageUrl = ConfigurationManager.AppSettings["MasterPageUrl"];
        if (string.IsNullOrEmpty(MasterPageUrl))
            MasterPageUrl = "~/_layouts/DarMasterPages/PWSv4.master";
        this.MasterPageFile = MasterPageUrl;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        string FormName = "";
        string ListId = Request.QueryString["ListId"];
        string ItemId = Request.QueryString["ID"];


        string Reference = "",
               EmailSubject = "",
               EmailBody = "";

        SPWeb web = null;
        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    web = site.OpenWeb();
                }
            });
            web.AllowUnsafeUpdates = true;

            Guid newGuid = new Guid(ListId);
            SPList list = web.Lists[newGuid];
            SPListItem item = list.GetItemById(int.Parse(ItemId));

            string targetColumns = "ReviewedURL";
            if (item[targetColumns] != null && !string.IsNullOrEmpty(item[targetColumns].ToString()))
            {
                SPFieldUrlValue FileUrl = new SPFieldUrlValue(item[targetColumns].ToString());

                if (FileUrl != null && !string.IsNullOrEmpty(FileUrl.Url))
                {
                    SPFile file = web.GetFile(FileUrl.Url);
                    if (file.Exists)
                    {
                        #region PRINT PDF VERSION
                        byte[] filebyte = file.OpenBinary();
                        Response.Clear();
                        Response.ContentType = "application/pdf";
                        Response.AddHeader("Content-Length", filebyte.Length.ToString());
                        Response.AddHeader("Content-Disposition", "attachment;filename=" + file.Name);
                        Response.Buffer = true;
                        Response.OutputStream.Write(filebyte, 0, filebyte.Length);
                        Response.Flush();
                        Response.End();
                        #endregion
                    }
                }
            }

            else
            {
                string StrFullRef = "";
                if (item.ParentList.Fields.ContainsField("FullRef"))
                    StrFullRef = (item["FullRef"] != null) ? item["FullRef"].ToString() : "";

                string DeliverableType = (item["DeliverableType"] != null) ? item["DeliverableType"].ToString() : "";
                string ListName = "";
                SPListItem LeadItem = null;

                ListName = "Lead Action";

                if (DeliverableType == "MIR")
                    FormName = "MIR_FORM";


                else if (DeliverableType == "SI")
                {
                    FormName = "SI_FORM";
                    if (string.IsNullOrEmpty(StrFullRef))
                    {
                        string _Ref = (item["Reference"] != null) ? item["Reference"].ToString() : "";
                        StrFullRef = _Ref;
                    }
                }
                else if (DeliverableType == "DPR")
                {
                    FormName = "DPR_FORM";
                    if (string.IsNullOrEmpty(StrFullRef))
                    {
                        string _Ref = (item["Reference"] != null) ? item["Reference"].ToString() : "";
                        StrFullRef = _Ref;
                    }
                }
                else if (DeliverableType == "MAT")
                    FormName = "MAT_FORM";


                else if (DeliverableType == "SCR")
                    FormName = "SCR_FORM";

                else
                {
                    string BldgType = (item["InspType"] != null) ? item["InspType"].ToString().Split('#')[1] : "";
                    if (BldgType == "Buildings")
                        FormName = "RIW_BLDG";
                    else FormName = "OTHER_RIW";
                }

                if (DeliverableType != "SI" && DeliverableType != "DPR")
                {
                    SPList objlistLeadTasks = web.Lists[ListName];
                    SPQuery _Leadquery = new SPQuery();
                    _Leadquery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
                    SPListItemCollection LeadItems = objlistLeadTasks.GetItems(_Leadquery);
                    if (LeadItems.Count == 1)
                        LeadItem = LeadItems[0];
                }

                GenericEmailTemplate(web, FormName, item, ref EmailBody, ref EmailSubject, LeadItem, DeliverableType);
                Formlbl.Text = EmailBody;

                #region PRINT PDF VERSION
                byte[] CRSFrom_PDFByte = PDFConvert(Formlbl.Text);
                string attachName = StrFullRef + ".pdf";
                //if (item.Attachments.Count > 0)
                //{
                //    try { item.Attachments.DeleteNow(attachName); } catch { }
                //}
                //item.Attachments.Add(attachName, CRSFrom_PDFByte); //File.ReadAllBytes(strpath3)
                //item.SystemUpdate(false);

                Response.Clear();
                Response.ContentType = "application/pdf";
                Response.AddHeader("Content-Length", CRSFrom_PDFByte.Length.ToString());
                Response.AddHeader("Content-Disposition", "attachment;filename=" + attachName);
                Response.Buffer = true;
                Response.OutputStream.Write(CRSFrom_PDFByte, 0, CRSFrom_PDFByte.Length);
                Response.Flush();
                Response.End();
                #endregion
            }
        }
        catch (ThreadAbortException) { }
        catch (Exception ex)
        {
            SendErrorEmail("SiteUrl: <br/>" + web.Url + "<br/><br/>" +
                           "Generic PrintView.aspx:<br/> Error on Page_Load Function <br/><br/> " +
                           "Reference: " + Reference + "<br/><br/>" +
                           "Message: <br/>" + ex.Message + "<br/><br/>" +
                           "StackTrace: <br/>" + ex.StackTrace);
        }
        finally
        {
            if (web != null)
                web.Dispose();
        }
    }

    public void GenericEmailTemplate(SPWeb web, string EmailName, SPListItem Listitem, ref string EmailBody, ref string EmailSubject, SPListItem LeadItem, string DeliverableType)
    {
        #region Extract Email Body, Subject, ListName From Emails List
        SPList Emails = web.Lists["Emails"];
        SPQuery EmailQuery = new SPQuery();
        EmailQuery.Query = "<Where>" +
                               "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + EmailName + "</Value></Eq>" +
                           "</Where>";
        SPListItemCollection items = Emails.GetItems(EmailQuery);
        SPListItem item = items[0];
        EmailSubject = (item["EmailSubject"] != null) ? item["EmailSubject"].ToString() : "";
        EmailBody = (item["Body"] != null) ? item["Body"].ToString() : "";

        if(DeliverableType == "SI")
        {
            int Count = 0; 
            string Response = GetVersionedMultiLineTextAsPlainText(Listitem, "Response", ref Count, false, false);

            if (Count > 1)
            {
                #region APPEND RESPONSE APPENDIX FOR SITE INSTRUCTION FORM
                EmailBody += "<table style='border-collapse: collapse; border: black 0px solid; color:black; margin:auto;' border='0' cellspacing='0' cellpadding='2' width='100%' align='center'>" +
                            "<tbody>" +
                                "<tr><td><br /><br /></td></tr>" +
                            "<tr height='20'>" +
                            "<td style='border: black 1px solid; border-bottom: 0;' width='30%' align='center'>" +
                            "<p><strong>Employer:</strong></p>" +
                            "<p><strong>[EmployerImg]</strong></p>" +
                            "</td>" +
                            "<td style='border: black 1px solid; border-bottom: 0;' colspan='2' align='center'><strong>Engineer:<br /></strong>" +
                            "<p><strong>[DarImg]</strong></p>" +
                            "</td>" +
                            "<td style='border: black 1px solid; border-bottom: 0;' width='30%' align='center'><strong>Contractor:<br /></strong>" +
                            "<p><strong>[ContractorImg]</strong></p>" +
                            "</td>" +
                            "</tr>" +
                            "<tr>" +
                            "<td style='border: black 1px solid; border-top: 0;' width='30%' align='center'>&nbsp;</td>" +
                            "<td style='border: black 1px solid; border-top: 0;' colspan='2' align='center'>&nbsp;</td>" +
                            "<td style='border: black 1px solid; border-top: 0;' width='30%' align='center'>&nbsp;</td>" +
                            "</tr>" +
                            "<tr>" +
                            "<td>&nbsp;</td>" +
                            "<td colspan='2'>&nbsp;</td>" +
                            "<td>&nbsp;</td>" +
                            "</tr>" +
                            "<tr>" +
                            "<td style='text-align: center; border: black 1px solid;' colspan='4'><span style='font-size: medium; font-family: arial,helvetica,sans-serif;'>" +
                            "<p><strong>E-Site Instruction</strong></p>" +
                            "</span></td>" +
                            "</tr>" +
                            "<tr>" +
                            "<td>&nbsp;</td>" +
                            "<td colspan='2'>&nbsp;</td>" +
                            "<td>&nbsp;</td>" +
                            "</tr>" +

                            "<tr>" +
                            "<td style='font-size:20px;' colspan='4'><p><strong><u>Response Appendix:</u></strong></p></td>" +
                            "</tr>" +

                            "<tr>" +
                            "<td colspan='4'></td>" +
                            "</tr>" +

                            "<tr>" +
                            "<td colspan='4'><p>[AppResponse]</p></td>" +
                            "</tr>" +
                            "</tbody>" +
                            "</table>";
                #endregion
            }
        }

        string ListName = (item["ListName"] != null) ? item["ListName"].ToString() : "";
        Dictionary<string, string> EmailData = new Dictionary<string, string>();
        EmailData.Add("Subject", EmailSubject);
        EmailData.Add("Body", EmailBody);
        #endregion

        string Code = "";
        if (DeliverableType == "DPR")
            Code = (Listitem["Status"] != null) ? Listitem["Status"].ToString() : "";
        else Code = (Listitem["Code"] != null) ? Listitem["Code"].ToString() : "";
        string WF_Status = "Open";

        if (Code != "Open" && !string.IsNullOrEmpty(Code))
            WF_Status = "Closed";

        string LeadName = "";
        if (LeadItem != null)
        {
            SPFieldUserValueCollection LeadUser = (LeadItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(web, LeadItem["AssignedTo"].ToString()) : null;
            if (LeadUser != null && LeadUser.Count > 0)
                LeadName = LeadUser[0].User.Name;
        }
       
        SPList List = web.Lists[ListName];
        List<SPListItem> ListItems = new List<SPListItem>();
        if (!ListItems.Contains(Listitem))
            ListItems.Add(Listitem);

        if (ListItems.Count > 0)
        {
            Regex re;
            string regularExpressionPattern = "",
                   inputText = "";
            bool SubjectVisited = false;

            string rejectedcomments = "";
            int counter = 0;
            int size = items.Count;

            foreach (SPListItem RLODitem in ListItems)
            {
                foreach (var KeyVal in EmailData)
                {
                    #region Fill InputText (Body, Subject)
                    string Key = KeyVal.Key;
                    regularExpressionPattern = @"\[(.*?)\]";
                    inputText = "";

                    if (Key == "Subject" && SubjectVisited == false)
                    {
                        inputText = EmailSubject;
                        SubjectVisited = true;
                    }
                    else
                        inputText = EmailBody;

                    re = new Regex(regularExpressionPattern);
                    #endregion

                    foreach (Match m in re.Matches(inputText))
                    {
                        ColumnInBrackets = m.Value;
                        string ColumnName = "";
                        ColumnName = ColumnInBrackets.Replace("[", "").Replace("]", "");
                        value = "";

                        if (ColumnInBrackets.Contains("|"))
                        {
                            #region if Column Name Format  = ListName|ColumnName
                            string[] SplitColumnInBrackets = ColumnInBrackets.Split('|');
                            string ColumnListName = "";
                            ColumnListName = SplitColumnInBrackets[0].Replace("[", "");
                            ColumnName = SplitColumnInBrackets[1].Replace("]", "");

                            #region Generate Query
                            SPQuery GenerateQuery = new SPQuery();

                            if (EmailType == "SingleItem")
                            {
                                GenerateQuery.Query = "<Where>" +
                                                      "<Eq><FieldRef Name='ID' /><Value Type='Counter'>" + int.Parse(EmailValue) + "</Value></Eq>" +
                                                    "</Where>";
                            }
                            else if (ColumnListName == GetParameter("LeadTasks") || ColumnListName == GetParameter("PartTasks"))
                            {
                                GenerateQuery.Query = "<Where>" +
                                                    "<Eq><FieldRef Name='Reference' /><Value Type='Text'>" + EmailValue + "</Value></Eq>" +
                                                  "</Where>";
                            }
                            else
                            {
                                GenerateQuery.Query = "<Where>" +
                                                    "<Eq><FieldRef Name='SubmittalRef' /><Value Type='Text'>" + EmailValue + "</Value></Eq>" +
                                                  "</Where>";
                            }
                            #endregion

                            List = web.Lists[ColumnListName];
                            SPListItemCollection Selecteditems = List.GetItems(GenerateQuery);
                            SPListItem Selecteditem = Selecteditems[0];

                            #region Replace ViewItem, EditItem, ViewAll by itemUrl
                            //[{ViewItem}] | [{EditItem}] | [{ViewAll}]</
                            if (ColumnName.ToLower() == "{viewitem}")
                            {
                                value = "<a href ='" + Selecteditem.Web.Url + "/" + Selecteditem.ParentList.RootFolder.Url + "/DispForm.aspx?ID=" + Selecteditem.ID + "'>View Item</a>";
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);

                                continue;
                            }
                            else
                                if (ColumnName.ToLower() == "{edititem}")
                            {
                                value = "<a href ='" + Selecteditem.Web.Url + "/" + Selecteditem.ParentList.RootFolder.Url + "/EditForm.aspx?ID=" + Selecteditem.ID + "'>Edit Item</a>";
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                continue;
                            }
                            else if (ColumnName.ToLower() == "{viewall}")
                            {
                                value = "<a href ='" + Selecteditem.Web.Url + "/" + Selecteditem.ParentList.RootFolder.Url + "/" + List.DefaultView + ".aspx'>View All</a>";
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                continue;
                            }
                            else if (ColumnName.ToLower() == "{submission-repository-link}")
                            {
                                //"/" + EmailValue +
                                value = "<a href ='" + destinationUrl(null, Selecteditem, true) + "'>Repository Folder</a>";
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                continue;
                            }
                            #endregion

                            #region Extract Body and Subject based on Column Type
                            string Display_Name = List.Fields.GetField(ColumnName).Title;
                            SPField field = List.Fields[Display_Name];
                            SPFieldType fieldType = field.Type;

                            if (Selecteditem[ColumnName] == null)
                            {
                                #region Set Empty String
                                value = "";
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, "");
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                                #endregion
                            }
                            else if (field.Type == SPFieldType.Text || field.Type == SPFieldType.Choice)
                            {
                                #region SPFieldType.Text OR field.Type == SPFieldType.Choice
                                value = (Selecteditem[ColumnName] != null) ? Selecteditem[ColumnName].ToString() : "";

                                if (Key == "Subject")
                                {
                                    if (ColumnName == "ProjectName")
                                    {
                                        value = GetParameter("ProjectName");
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    }
                                    else
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                }
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                #endregion
                            }
                            else
                                if (field.Type == SPFieldType.URL)
                            {
                                #region SPFieldType.URL
                                if (Selecteditem[ColumnName] != null)
                                {
                                    SPFieldUrlValue Link = new SPFieldUrlValue(Selecteditem[ColumnName].ToString());
                                    if (Link.Url != null)
                                    {
                                        if (Title == "Console_Email")
                                        {
                                            if (ColumnName == "ReviewLink")
                                                value = "<a href= '" + Link.Url + "'>Click Here</a>";
                                            else
                                                value = "<a href= '" + Link.Url + "'>" + Link.Description + "</a>";
                                        }
                                        else
                                            value = "<a href= '" + Link.Url + "'>" + Link.Description + "</a>";
                                    }
                                    else value = "";

                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            else
                                    if (field.Type == SPFieldType.DateTime)
                            {
                                #region SPFieldType.DateTime
                                if (Selecteditem[ColumnName] != null)
                                {
                                    value = ((DateTime)Selecteditem[ColumnName]).ToString("MMM dd, yyyy");
                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            else
                                        if (field.Type == SPFieldType.Lookup)
                            {
                                #region SPFieldType.Lookup
                                if (Selecteditem[ColumnName] != null)
                                {
                                    SPFieldLookupValue LookupColumn = (Selecteditem[ColumnName] != null) ? new SPFieldLookupValue(Selecteditem[ColumnName].ToString()) : null;
                                    if (LookupColumn != null)
                                        value = LookupColumn.LookupValue;

                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            else
                                            if (field.Type == SPFieldType.User)
                            {
                                #region SPFieldType.User
                                SPFieldUserValueCollection users = (Selecteditem[ColumnName] != null) ? new SPFieldUserValueCollection(web, Selecteditem[ColumnName].ToString()) : null;
                                if (users != null)
                                {
                                    if (users.Count != 0)
                                    {
                                        for (int i = 0; i < users.Count; i++)
                                        {
                                            if (!String.IsNullOrEmpty(users[i].User.Name))
                                            {
                                                value += users[i].User.Name + ",<br />";
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(value))
                                            value = value.Remove(value.Length - 1);
                                    }
                                }

                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                #endregion
                            }
                            else
                                                if (field.Type == SPFieldType.MultiChoice)
                            {
                                #region SPFieldType.User
                                if (Selecteditem[ColumnName] != null)
                                {
                                    SPFieldMultiChoiceValue itemValue = new SPFieldMultiChoiceValue(Selecteditem[ColumnName].ToString());
                                    for (int j = 0; j < itemValue.Count; j++)
                                    {
                                        value += itemValue[j] + ", ";
                                    }

                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            else if (field.Type == SPFieldType.Attachments)
                            {
                                #region SPFieldType.Attachments
                                for (int i = 0; i < Selecteditem.Attachments.Count; i++)
                                {
                                    if (value == "")
                                        value = "<a href='" + Selecteditem.Attachments.UrlPrefix + Selecteditem.Attachments[i] + "'>" + Selecteditem.Attachments[i] + "</a>";
                                    else
                                        value += "<br/><a href='" + Selecteditem.Attachments.UrlPrefix + Selecteditem.Attachments[i] + "'>" + Selecteditem.Attachments[i] + "</a>";
                                }
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                #endregion
                            }
                            #endregion

                            continue;
                            #endregion
                        }
                        else if (ColumnName == "ProjectName" && Key == "Subject")
                        {
                            #region if ProjectName Column
                            if (Key == "Subject")
                            {
                                value = GetParameter("ProjectName");
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                            }
                            continue;
                            #endregion
                        }
                        else if (ColumnName == "EmployerImg" || ColumnName == "DarImg" || ColumnName == "ContractorImg" || ColumnName.ToLower() == "contractorsignature")
                        {
                            #region if Logs Column
                            string ImgName = ColumnName, ContCode = "";

                            if (ColumnName == "ContractorImg" || ColumnName.ToLower() == "contractorsignature")
                            {
                                string isMultiContracotr = GetParameter("isMultiContracotr");
                                if (isMultiContracotr.ToLower() == "yes" || isMultiContracotr == "1")
                                {
                                    string tmp = Listitem.Url.Substring(Listitem.Url.ToLower().IndexOf("/lists/") + 7); //Strip /Lists/ and preceding chars
                                    string Listname = tmp.Substring(0, tmp.IndexOf("/"));
                                    tmp = tmp.Substring(tmp.IndexOf("/") + 1); //Strip "Listname/"
                                    if (tmp.Contains("/"))
                                        ContCode = tmp.Substring(0, tmp.IndexOf("/"));
                                    ImgName = ContCode + ImgName;
                                }
                            }

                            SPList _AdminPages = web.Lists["Images"];
                            SPQuery Logo = new SPQuery();
                            Logo.Query = "<Where><Contains><FieldRef Name='FileLeafRef' /><Value Type='Text'>" + ImgName + "</Value></Contains></Where>";
                            SPListItemCollection DarItemLogo = _AdminPages.GetItems(Logo);
                            if (DarItemLogo.Count > 0)
                            {
                                SPFile file = DarItemLogo[0].File;
                                string Url = web.Url + "/" + file.Url;
                                byte[] inputImageStream = web.GetFile(Url).OpenBinary();

                                string Extension = Path.GetExtension(file.Url).Replace(".", "");
                                value = "data:image/" + Extension + ";base64," + Convert.ToBase64String(inputImageStream);
                                value = "<img alt='' src='" + value + "' height='50' />";
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                            #endregion
                        }

                        else if (ColumnName == "SiteStaff" || ColumnName == "DelivMat" || ColumnName == "SiteEquipment" || ColumnName == "SiteVisitors" || ColumnName == "WorksDone" || ColumnName == "SiteRemarks")
                        {
                            int DefaultRows = 16;
                            string Listname = "", HeaderTitle = "";
                            string[] Columns = { "Title", "No", "Hrs" };
                            string[] Width = { "75%", "10%", "10%" };

                            if (ColumnName == "SiteStaff")
                            {
                                Listname = "On Site Staff and Labour";
                                HeaderTitle = "ON SITE STAFF & LABOUR RECORD";
                            }
                            else if (ColumnName == "DelivMat")
                            {
                                Listname = "Delivery of Materials";
                                HeaderTitle = "DELIVERY OF MATERIALS";
                                Columns = new[] { "Title", "Qty", "Unit" };
                            }
                            else if (ColumnName == "SiteEquipment")
                            {
                                Listname = "On Site Plant & Equipment Record";
                                HeaderTitle = "ON SITE PLANT & EQUIPMENT RECORD";
                            }
                            else if (ColumnName == "SiteVisitors")
                            {
                                Listname = "Visitors on Site";
                                HeaderTitle = "VISITORS ON SITE";
                                Columns = new[] { "Title", "Organization" };
                                Width = new[] { "55%", "40%" };
                                DefaultRows = 11;
                            }
                            else if (ColumnName == "WorksDone")
                            {
                                Listname = "Description of Works Done";
                                HeaderTitle = "DESCRIPTION OF WORKS DONE";
                                Columns = new[] { "Title", "Qty" };
                                Width = new[] { "80%", "10%" };
                                DefaultRows = 11;
                            }
                            else if (ColumnName == "SiteRemarks")
                            {
                                Listname = "DPR Remarks";
                                HeaderTitle = "REMARKS";
                                Columns = new[] { "Title" };
                                Width = new[] { "90%" };
                                DefaultRows = 11;
                            }

                            SPList list = web.Lists[Listname];
                            value = SET_HTML_TABLE(Listitem.ID, list, HeaderTitle, Columns, Width, DefaultRows);
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }

                        else if (ColumnName == "Esign")
                        {
                            if (WF_Status == "Closed")
                            {
                                #region if Digital Signature Column
                                SPList _AdminPages = web.Lists["ESignatures"];
                                SPQuery Logo = new SPQuery();
                                Logo.Query = "<Where><Eq><FieldRef Name='User' /><Value Type='User'>" + LeadName + "</Value></Eq></Where>";
                                SPListItemCollection DarItemLogo = _AdminPages.GetItems(Logo);
                                if (DarItemLogo.Count > 0)
                                {
                                    SPFile file = DarItemLogo[0].File;
                                    string Url = web.Url + "/" + file.Url;
                                    byte[] inputImageStream = web.GetFile(Url).OpenBinary();

                                    string Extension = Path.GetExtension(file.Url).Replace(".", "");
                                    value = "data:image/" + Extension + ";base64," + Convert.ToBase64String(inputImageStream);
                                    value = "<img alt='' src='" + value + "' width='150' height='35' align ='middle' />";
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                                #endregion
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName == "REEngComment")
                        {
                            if (WF_Status == "Closed")
                            {
                                value = (LeadItem["Comment"] != null) ? LeadItem["Comment"].ToString() : "";
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName == "RECode")
                        {
                            if (WF_Status == "Closed")
                            {
                                bool isInput = false;
                                value = (Listitem["Code"] != null) ? Listitem["Code"].ToString() : "";
                                SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "");

                                if (!isInput)
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName == "SentToContractorDate")
                        {
                            if (WF_Status == "Closed")
                            {
                                #region if E-Inspection Date Closed
                                if (Listitem[ColumnName] != null)
                                    value = ((DateTime)Listitem[ColumnName]).ToString("MMM dd, yyyy");
                                else value = DateTime.Today.ToString("MMM dd, yyyy");
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                #endregion
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName == "REName")
                        {
                            value = LeadName;
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName == "RE")
                        {
                            value = "";
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }


                        else if (ColumnName == "ProjTitle")
                        {
                            value = GetParameter("ProjectTitle");
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName == "PurposeChoices")
                        {
                            SetMultiChoiceLookups(web, Listitem, ref EmailBody, false);
                        }
                        else if (ColumnName == "DescriptionChoices")
                        {
                            SetMultiChoiceLookups(web, Listitem, ref EmailBody, true);
                        }
                        else if (ColumnName == "InspectionReport")
                        {
                            if (WF_Status == "Closed")
                            {
                                #region SET INSPECTION REVIEWERS TABLE
                                SPList list = web.Lists.TryGetList("Inspection Part Tasks");
                                if (list != null)
                                {
                                    SPQuery listQuery = new SPQuery();
                                    listQuery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + Listitem["FullRef"].ToString() + "</Value></Eq></Where>";
                                    SPListItemCollection _items = list.GetItems(listQuery);

                                    if (_items.Count > 0)
                                    {
                                        value = "<table style='border-collapse:collapse; color:black;'  width='100%' cellpadding='0' cellspacing='0'>" +
                                                "<tr>";
                                        value += "<td width='15%' style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>Trade</td>" +
                                                  "<td width='20%' style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>Insp. Name</td>" +
                                                  "<td width='55%' style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>Reply</td>" +
                                                  "<td width='10%' style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>Closed Date</td>" +
                                                 "</tr>";

                                        foreach (SPListItem listItem in _items)
                                        {
                                            string Trade = (listItem["Trade"] != null) ? listItem["Trade"].ToString() : "";
                                            string Editor = (listItem["Editor"] != null) ? listItem["Editor"].ToString().Split('#')[1] : "";
                                            string Comment = (listItem["Comment"] != null) ? listItem["Comment"].ToString() : "";
                                            string Date_Closed = (listItem["Date_Closed"] != null) ? ((DateTime)listItem["Date_Closed"]).ToString("MMM dd, yyyy") : "";

                                            value += "<tr><td style='border: 1px solid Black;'>" + Trade + "</td>" +
                                                         "<td style='border: 1px solid Black;'>" + Editor + "</td>" +
                                                         "<td style='border: 1px solid Black;'>" + Comment + "</td>" +
                                                         "<td style='border: 1px solid Black;' align='center'>" + Date_Closed + "</td></tr>";
                                        }

                                        value += "</table>";
                                    }
                                }
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                #endregion
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName == "MIRInspectionReport")
                        {
                            #region MIR INSPECTION REPORT COLUMN
                            value = "";

                            SPList list = web.Lists.TryGetList("MIR Part Tasks");
                            if (list != null)
                            {
                                SPQuery listQuery = new SPQuery();
                                listQuery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + Listitem["FullRef"].ToString() + "</Value></Eq></Where>";
                                SPListItemCollection _items = list.GetItems(listQuery);

                                if (_items.Count > 0)
                                {
                                    SPListItem _item = _items[0];
                                    string HandledBy = (_item["HandledBy"] != null) ? _item["HandledBy"].ToString() : "";
                                    bool ph_dmg = (_item["ph_dmg"] != null) ? bool.Parse(_item["ph_dmg"].ToString()) : false;
                                    bool det_corr = (_item["det_corr"] != null) ? bool.Parse(_item["det_corr"].ToString()) : false;
                                    bool app_sub = (_item["app_sub"] != null) ? bool.Parse(_item["app_sub"].ToString()) : false;
                                    bool acc_spl = (_item["acc_spl"] != null) ? bool.Parse(_item["acc_spl"].ToString()) : false;

                                    value = "<table align='left' style='border-collapse:collapse; color:black;' width='100%' cellpadding='10' cellspacing='0'>" +
                                                "<tr>" +
                                                    "<td>" +
                                                     "<table style='border-collapse:collapse; color:black' width='60%' cellpadding='0' cellspacing='0'>" +
                                                         "<tr>" +
                                                             "<td style='text -align:center; vertical-align:middle; background-color:#C0C0C0; font-weight:bold; border: black 1px solid;'>Question</td>" +
                                                             "<td style='text -align:center; vertical-align:middle; background-color:#C0C0C0; font-weight:bold; border: black 1px solid;'>Yes/No</td>" +
                                                             "<td style='text -align:center; vertical-align:middle; background-color:#C0C0C0; font-weight:bold; border: black 1px solid;'>Inspector Name</td>" +
                                                         "</tr>" +

                                                         "<tr>" +
                                                             "<td style='border: black 1px solid;'>Physical damage ?</td>" +
                                                              "<td style='text -align:center; border: black 1px solid;'>";

                                    if (ph_dmg)
                                        value += "<input name='1111' type='checkbox' checked />Yes &nbsp; <input name='2222' type='checkbox' />No</td>";
                                    else value += "<input name='1111' type='checkbox' />Yes &nbsp; <input name='2222' type='checkbox' checked />No</td>";

                                    value += "<td rowspan='4' style='text-align:center; border: black 1px solid;'>" + HandledBy + "</td>" +
                                         "</tr>" +

                                            "<tr>" +
                                                "<td style='border: black 1px solid; '>Details given above correct ?</td>" +
                                                "<td style='text -align:center; border: black 1px solid;'>";

                                    if (det_corr)
                                        value += "<input name='1111' type='checkbox' checked />Yes &nbsp; <input name='2222' type='checkbox' />No</td>";
                                    else value += "<input name='1111' type='checkbox' />Yes &nbsp; <input name='2222' type='checkbox' checked />No</td>";

                                    value += "</tr>" +
                                                "<tr>" +
                                                    "<td style='border: black 1px solid;'>Conform with approved submission ?</td>" +
                                                    "<td style='text -align:center; border: black 1px solid;'>";

                                    if (app_sub)
                                        value += "<input name='1111' type='checkbox' checked />Yes &nbsp; <input name='2222' type='checkbox' />No</td>";
                                    else value += "<input name='1111' type='checkbox' />Yes &nbsp; <input name='2222' type='checkbox' checked />No</td>";

                                    value += "</tr>" +
                                                "<tr>" +
                                                    "<td style='border: black 1px solid;'>Accessories/ spare parts included ?</td>" +
                                                    "<td style='text -align:center; border: black 1px solid;'>";

                                    if (acc_spl)
                                        value += "<input name='1111' type='checkbox' checked />Yes &nbsp; <input name='2222' type='checkbox' />No</td>";
                                    else value += "<input name='1111' type='checkbox' />Yes &nbsp; <input name='2222' type='checkbox' checked />No</td>";

                                    value += "</tr>" +
                                          "</table>" +
                                         "</td>" +
                                       "</tr>" +
                                     "</table>";
                                }
                            }


                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                            #endregion
                        }
                        else if (ColumnName == "AppResponse")
                        {
                            int Count = 0;
                            value = GetVersionedMultiLineTextAsPlainText(RLODitem, "Response", ref Count, false, true);

                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName == "LatestResponse")
                        {
                            int Count = 0;
                            value = GetVersionedMultiLineTextAsPlainText(RLODitem, "Response", ref Count, true, false);

                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName.Contains("/Time"))
                        {
                            value = ((DateTime)RLODitem[ColumnName.Split('/')[0]]).ToString("hh:mm tt");

                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }


                        else if (ColumnName.ToLower() == "currentuser")
                        {
                            if (WF_Status == "Closed")
                            {
                                #region if CurrentUser Column
                                value = LeadName; //SPContext.Current.Web.CurrentUser.Name;
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                {
                                    if (isMulti)
                                        ConcatEditedChunk();
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            else EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                            continue;
                        }
                        else if (ColumnName.ToLower() == "addedparttrade")
                        {
                            #region if AddedPartTrade Column For Reassign Form

                            value = "AddedPartTrade";
                            if (Key == "Subject")
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                            else
                            {
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            continue;
                            #endregion
                        }
                        else if (ColumnName.ToLower() == "lod-path")
                        {
                            #region if LOD_PATH Column For Reassign Form

                            string TentativeListName = GetParameter("TentativeLOD");

                            value = "<a href='" + web.Url + "/" + TentativeListName + "'>LOD Link</a>";

                            if (Key == "Subject")
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                            else
                            {
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            continue;
                            #endregion
                        }
                        else if (ColumnName.ToLower() == "weburl")
                        {
                            #region if WebUrl
                            value = web.Url;

                            if (Key == "Subject")
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                            else
                            {
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            continue;
                            #endregion
                        }
                        else if (ColumnName.ToLower() == "filterview")
                        {
                            #region if filterview
                            string DarTrade = (RLODitem["DarTrade"] != null) ? RLODitem["DarTrade"].ToString() : "";
                            value = "<a href='" + web.Url + "/Lists/" + List.Title + "/AllItems.aspx?FilterField1=DarTrade&FilterValue1=" + DarTrade + "'>View Items</a>";

                            if (Key == "Subject")
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                            else
                            {
                                if (isMulti)
                                    ConcatEditedChunk();
                                else
                                    EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            }
                            continue;
                            #endregion
                        }

                        #region Replace ViewItem, EditItem, ViewAll by itemUrl
                        //[{ViewItem}] | [{EditItem}] | [{ViewAll}]</
                        else if (ColumnName.ToLower() == "{viewitem}")
                        {
                            value = "<a href ='" + RLODitem.Web.Url + "/" + RLODitem.ParentList.RootFolder.Url + "/DispForm.aspx?ID=" + RLODitem.ID + "'>View Item</a>";
                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);

                            continue;
                        }
                        else
                            if (ColumnName.ToLower() == "{edititem}")
                        {
                            value = "<a href ='" + RLODitem.Web.Url + "/" + RLODitem.ParentList.RootFolder.Url + "/EditForm.aspx?ID=" + RLODitem.ID + "'>Edit Item</a>";
                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName.ToLower() == "{viewall}")
                        {
                            value = "<a href ='" + RLODitem.Web.Url + "/" + RLODitem.ParentList.RootFolder.Url + "/" + List.DefaultView + ".aspx'>View All</a>";
                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName.ToLower() == "{submission-repository-link}")
                        {
                            value = "<a href ='" + destinationUrl(null, RLODitem, true) + "'>Repository Folder</a>";
                            if (isMulti)
                                ConcatEditedChunk();
                            else
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        #endregion

                        else
                        {
                            if (Key == "Subject" && ColumnName == "ProjectName")
                            {
                                #region Get ProjectName from Parameters
                                value = GetParameter("ProjectName");
                                EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                continue;
                                #endregion
                            }

                            else if (ColumnName.ToLower() == "currentuser")
                            {
                                #region if CurrentUser Column

                                value = SPContext.Current.Web.CurrentUser.Name;
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                {
                                    if (isMulti)
                                        ConcatEditedChunk();
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                continue;
                                #endregion
                            }

                            string Display_Name = List.Fields.GetField(ColumnName).Title;

                            SPField field = List.Fields[Display_Name];
                            SPFieldType fieldType = field.Type;

                            #region CHECK if Body is Multi (id=items)
                            if (Key == "Body")
                            {
                                if (EmailBody.Contains("<tr id=\"items\">"))
                                {
                                    isMulti = true;

                                    if (isFirstVisit == false)
                                    {
                                        #region Multi Entry
                                        HeaderChunk = EmailBody.Substring(0, EmailBody.IndexOf("<tr id=\"items\">"));
                                        EditedChunk = EmailBody.Substring(EmailBody.IndexOf("<tr id=\"items\">"));
                                        FooterChunck = EditedChunk.Substring(EditedChunk.IndexOf("</tr>")).Replace("</tr>", "");
                                        EditedChunk = EditedChunk.Substring(0, EditedChunk.IndexOf("</table>")).Replace("</tbody>", "").Replace("</table>", "");
                                        AddNewRow = EditedChunk;
                                        isFirstVisit = true;
                                        #endregion
                                    }
                                }
                            }
                            #endregion

                            #region Extract Body and Subject based on Column Type
                            if (RLODitem[ColumnName] == null)
                            {
                                #region Set Empty String
                                value = "";
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, "");
                                else
                                {
                                    if (isMulti)
                                        ConcatEditedChunk();
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, "");
                                }
                                #endregion
                            }

                            else if (field.Type == SPFieldType.Text || field.Type == SPFieldType.Note)
                            {
                                #region SPFieldType.Text OR field.Type == SPFieldType.Choice

                                value = (RLODitem[ColumnName] != null) ? RLODitem[ColumnName].ToString() : "";

                                if (ColumnName.ToLower().Equals("rejectedcomments"))
                                {
                                    #region  Rejected Comments
                                    if (isMulti)
                                    {
                                        List<string> temp = new List<string>();

                                        temp = Regex.Split(value, "<hr />").ToList();

                                        value = temp[temp.Count - 2];

                                        if (!rejectedcomments.Contains(value))
                                            rejectedcomments += value + "<hr>";

                                        if (counter == size - 1) //last item
                                        {
                                            HeaderChunk = HeaderChunk.Replace(ColumnInBrackets, rejectedcomments);

                                            FooterChunck = FooterChunck.Replace(ColumnInBrackets, rejectedcomments);

                                            if (isNewRow)
                                            {
                                                EditedChunk = EditedChunk.Replace(ColumnInBrackets, rejectedcomments) + AddNewRow.Replace(ColumnInBrackets, rejectedcomments);
                                                isNewRow = false;
                                            }
                                            else
                                            {
                                                EditedChunk = EditedChunk.Replace(ColumnInBrackets, rejectedcomments);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        List<string> temp = new List<string>();

                                        temp = Regex.Split(value, "<hr>").ToList();

                                        value = temp[temp.Count - 2];

                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                    #endregion
                                }
                                else
                                {
                                    if (Key == "Subject")
                                    {
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    }
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else if (field.Type == SPFieldType.Choice || field.Type == SPFieldType.MultiChoice || field.Type == SPFieldType.Boolean)
                            {
                                #region CHOICE, MULTI-CHOICE, BOOLEAN
                                bool isInput = false;
                                if (field.Type == SPFieldType.Boolean)
                                {
                                    value = (RLODitem[ColumnName] != null) ? RLODitem[ColumnName].ToString() : "False";
                                    SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "Boolean");
                                }

                                else if (field.Type == SPFieldType.Choice)
                                {
                                    value = (RLODitem[ColumnName] != null) ? RLODitem[ColumnName].ToString() : "";
                                    SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "");
                                }

                                else
                                {
                                    SPFieldMultiChoiceValue ListValues = (RLODitem[ColumnName] != null) ? new SPFieldMultiChoiceValue(RLODitem[ColumnName].ToString()) : null;
                                    if (ListValues != null)
                                    {
                                        string ConcatValues = "";
                                        for (int i = 0; i < ListValues.Count; i++)
                                        {
                                            value = ListValues[i].ToString();
                                            SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "");
                                            ConcatValues += value + ",";
                                        }
                                        if (!string.IsNullOrEmpty(ConcatValues))
                                        {
                                            ConcatValues = ConcatValues.TrimEnd(',');
                                            value = ConcatValues;
                                        }
                                    }
                                }

                                if (!isInput)
                                {
                                    if (Key == "Subject")
                                    {
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    }
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else
                                if (field.Type == SPFieldType.URL)
                            {
                                #region SPFieldType.URL
                                if (RLODitem[ColumnName] != null)
                                {
                                    SPFieldUrlValue Link = new SPFieldUrlValue(RLODitem[ColumnName].ToString());
                                    if (Link.Url != null)
                                    {
                                        if (Title == "Console_Email")
                                        {
                                            if (ColumnName == "ReviewLink")
                                                value = "<a href= '" + Link.Url + "'>Click Here</a>";
                                            else
                                                value = "<a href= '" + Link.Url + "'>" + Link.Description + "</a>";
                                        }
                                        else
                                            value = "<a href= '" + Link.Url + "'>" + Link.Description + "</a>";
                                    }
                                    else value = "";

                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else
                              if (field.Type == SPFieldType.DateTime)
                            {
                                #region SPFieldType.DateTime

                                if (RLODitem[ColumnName] != null)
                                {
                                    value = ((DateTime)RLODitem[ColumnName]).ToString("MMM dd, yyyy");
                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else
                             if (field.Type == SPFieldType.Lookup)
                            {
                                #region SPFieldType.Lookup
                                value = "";
                                bool isInput = false;
                                SPFieldLookupValueCollection LookupColumn = (RLODitem[ColumnName] != null) ? new SPFieldLookupValueCollection(RLODitem[ColumnName].ToString()) : null;
                                if (LookupColumn != null)
                                {
                                    string ConcatValues = "";
                                    foreach (SPFieldLookupValue val in LookupColumn)
                                    {
                                        value = val.LookupValue;
                                        SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "");
                                        ConcatValues += value + ",";
                                    }
                                    if (!string.IsNullOrEmpty(ConcatValues))
                                    {
                                        ConcatValues = ConcatValues.TrimEnd(',');
                                        value = ConcatValues;
                                    }
                                }

                                if (!isInput)
                                {
                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else
                             if (field.Type == SPFieldType.User)
                            {
                                #region SPFieldType.User
                                SPFieldUserValueCollection users = (RLODitem[ColumnName] != null) ? new SPFieldUserValueCollection(web, RLODitem[ColumnName].ToString()) : null;
                                if (users != null)
                                {
                                    if (users.Count != 0)
                                    {
                                        for (int i = 0; i < users.Count; i++)
                                        {
                                            if (!String.IsNullOrEmpty(users[i].User.Name))
                                            {
                                                value += users[i].User.Name + ",<br />";
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(value))
                                            value = value.Remove(value.Length - 7);
                                    }
                                }

                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                {
                                    if (isMulti)
                                        ConcatEditedChunk();
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }

                                #endregion
                            }
                            else
                             if (field.Type == SPFieldType.MultiChoice)
                            {
                                #region SPFieldType.MultiChoice
                                if (RLODitem[ColumnName] != null)
                                {
                                    SPFieldMultiChoiceValue itemValue = new SPFieldMultiChoiceValue(RLODitem[ColumnName].ToString());
                                    for (int j = 0; j < itemValue.Count; j++)
                                    {
                                        value += itemValue[j] + ",";
                                    }

                                    if (!string.IsNullOrEmpty(value))
                                        value = value.TrimEnd(',');

                                    if (Key == "Subject")
                                        EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                    else
                                    {
                                        if (isMulti)
                                            ConcatEditedChunk();
                                        else
                                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                    }
                                }
                                #endregion
                            }
                            else
                              if (field.Type == SPFieldType.Attachments)
                            {
                                #region SPFieldType.Attachments
                                for (int i = 0; i < RLODitem.Attachments.Count; i++)
                                {
                                    if (value == "")
                                        value = "<a href='" + RLODitem.Attachments.UrlPrefix + RLODitem.Attachments[i] + "'>" + RLODitem.Attachments[i] + "</a>";
                                    else
                                        value += "<br/><a href='" + RLODitem.Attachments.UrlPrefix + RLODitem.Attachments[i] + "'>" + RLODitem.Attachments[i] + "</a>";
                                }
                                if (Key == "Subject")
                                    EmailSubject = EmailSubject.Replace(ColumnInBrackets, value);
                                else
                                {
                                    if (isMulti)
                                        ConcatEditedChunk();
                                    else
                                        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                                }
                                #endregion
                            }
                            #endregion
                        }
                    }
                } // End KeyVal
                isNewRow = true;

                counter++;
            }//end of foreach

            if (isMulti)
            {
                EmailBody = HeaderChunk + EditedChunk + FooterChunck;
            }
        }
    }

    public string ConcatEditedChunk()
    {
        HeaderChunk = HeaderChunk.Replace(ColumnInBrackets, value);
        FooterChunck = FooterChunck.Replace(ColumnInBrackets, value);

        if (isNewRow)
        {
            EditedChunk = EditedChunk.Replace(ColumnInBrackets, value) + AddNewRow.Replace(ColumnInBrackets, value);
            isNewRow = false;
        }
        else
        {
            EditedChunk = EditedChunk.Replace(ColumnInBrackets, value);
        }
        return EditedChunk;
    }

    public void SetInputTag(string inputText, string ColumnName, string ColumnValue, ref string EmailBody, ref bool isInput, string ColType)
    {

        Regex TagReg;
        if(ColType == "Boolean")
          TagReg = new Regex(@"<input[^>]*name\s*=\s*" + "\"" + @"\[" + ColumnName + @"[^>]*>");
        else TagReg = new Regex(@"<input[^>]*name\s*=\s*('|""|)\[" + ColumnName + @"[^>]*>(\s*" + ColumnValue + @"|\S*" + ColumnValue + ")");

        foreach (Match input in TagReg.Matches(inputText))
        {
            string TagPhrase = "";
            TagPhrase = input.Value;
            string FirstTagChunk = TagPhrase.Substring(0, TagPhrase.IndexOf('>'));
            string TagText = TagPhrase.Substring(TagPhrase.IndexOf('>') + 1);
            if (!string.IsNullOrEmpty(TagText))
                TagText = TagText.Trim();

            if (TagText.ToLower() == ColumnValue.ToLower() || ColumnValue.ToLower() == "true")
            {
                string AdjustedTag = FirstTagChunk + " checked />" + TagText;
                EmailBody = EmailBody.Replace(TagPhrase, AdjustedTag);
                isInput = true;
                return;
            }
        }
    }

    public string destinationUrl(SPListItem item, SPListItem RLODItem, bool GetFolderPath)
    {
        //string FullUrl = "";
        //string ConcateFolderStructure = "";

        //#region Get Values of DelivFolderStructure

        //string[] SplitDelivFolderStructure = DelivFolderStructure.Split('-');
        //for (int i = 0; i < SplitDelivFolderStructure.Length; i++)
        //{
        //    string ColumnName = SplitDelivFolderStructure[i];
        //    if (ColumnName == "Acronym")
        //    {
        //        ConcateFolderStructure += Acronym + "/";
        //    }
        //    else
        //    {
        //        ConcateFolderStructure += RLODItem[SplitDelivFolderStructure[i]].ToString() + "/";
        //    }
        //}

        //#endregion

        //FilesRepositoryUrl = FilesRepositoryUrl.Trim();
        //if (FilesRepositoryUrl[FilesRepositoryUrl.Length - 1] == '/') // Remove / from the end of Url if Exist
        //    FilesRepositoryUrl = FilesRepositoryUrl.Remove(FilesRepositoryUrl.Length - 1);

        //if (GetFolderPath)
        //{
        //    if (MutliRepository.ToLower() == "yes" || MutliRepository == "1")
        //        FullUrl = FilesRepositoryUrl + "/" + RLODItem[MultiRepositoryFieldName] + "/" + RepositoryListName + "/" + SubmittedDelivFolderName + "/" + ConcateFolderStructure;
        //    else
        //        FullUrl = FilesRepositoryUrl + "/" + RepositoryListName + "/" + SubmittedDelivFolderName + "/" + ConcateFolderStructure;
        //}
        //else
        //{
        //    if (MutliRepository.ToLower() == "yes" || MutliRepository == "1")
        //        FullUrl = FilesRepositoryUrl + "/" + RLODItem[MultiRepositoryFieldName] + "/" + RepositoryListName + "/" + SubmittedDelivFolderName + "/" + ConcateFolderStructure + item.File.Name;
        //    else
        //        FullUrl = FilesRepositoryUrl + "/" + RepositoryListName + "/" + SubmittedDelivFolderName + "/" + ConcateFolderStructure + item.File.Name;
        //}
        return "";
        //return FullUrl;
    }

    public string SET_HTML_TABLE(int ItemID, SPList list, string TableHeader, string[] Columns, string[] Width, int defaultRowNum)
    {
        string ViewFields = "";
        SPQuery _query = new SPQuery();
        _query.Query = "<Where><Eq><FieldRef Name='Lookup_ID' /><Value Type='Lookup'>" + ItemID + "</Value></Eq></Where>";

        string HtmlString = "<table style='border-collapse:collapse; color:black; font-size:12px'  width='100%' cellpadding='0' cellspacing='0'><tr>";

        HtmlString += "<tr><td style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center' colspan=" + (Columns.Length + 1) + ">" + TableHeader + "</td></tr>";
        HtmlString += "<td width='5%' style='border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>No.</td>";

        for (int i = 0; i < Columns.Length; i++)
        {
            string Field = Columns[i];
            ViewFields += "<FieldRef Name='" + Field + "' />";

            string DisplayName = list.Fields.GetField(Field).Title;
            HtmlString += "<td width=" + Width[i] + " style = 'border: 1px solid Black; font-weight: bold; background-color:#C0C0C0;' align='center'>" + DisplayName + "</td>";
        }
        HtmlString += "</tr>";

        _query.ViewFields = ViewFields;
        SPListItemCollection items = list.GetItems(_query);
        //if (items.Count > 0)
        //{
            int RowNum = 1;
            foreach (SPListItem item in items)
            {
                SET_HTML_TABLE_METAINFO(item, ref RowNum, ref HtmlString, Columns);
            }

            while (defaultRowNum > RowNum)
            {
                SET_HTML_TABLE_METAINFO(null, ref RowNum, ref HtmlString, Columns);
            }
        //}
        HtmlString += "</table>";
        return HtmlString;
    }

    public void SET_HTML_TABLE_METAINFO(SPListItem item, ref int RowNum, ref string HtmlString, string[] Columns)
    {
        HtmlString += "<tr>";
        HtmlString += "<td style='border: 1px solid Black;' align='center'>" + RowNum + "</td>";
        for (int j = 0; j < Columns.Length; j++)
        {
            string ColumnName = Columns[j];
            object columnValue = "";

            if (item != null)
                columnValue = item[ColumnName];

            HtmlString += "<td style='border: 1px solid Black;'";

            if (ColumnName == "No" || ColumnName == "Qty" || ColumnName == "Unit" || ColumnName == "Hrs")
                HtmlString += " align='center'> " + columnValue + "</td>";
            else HtmlString += "> &nbsp;" + columnValue + "</td>";
        }
        HtmlString += "</tr>";
        RowNum++;
    }

    public static string GetParameter(string key)
    {
        string retValue = "";
        SPWeb web = null;
        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                { web = site.OpenWeb(); }
            });

            web.AllowUnsafeUpdates = true;

            SPList list = web.Lists["Parameters"];
            SPQuery spquery = new SPQuery();
            spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                            key + "</Value></Eq></Where>";
            SPListItemCollection colitems = list.GetItems(spquery);
            if (colitems.Count == 1)
            {
                SPListItem item = colitems[0];
                if (item["Value"] != null && item["Value"].ToString() != "")
                    retValue = item["Value"].ToString().Trim();
            }
            else if (colitems.Count == 0)
            {
                if (key == "EnableCustomLinks" || key == "BIMRevFormat" || key == "Case_Sensative" || key == "ManageAllowedExtensions" || key == "Enable_Phase_Validation" || key == "Phase_Field" ||
                    key == "Enable_Xls_Creation" || key == "XlsColumnsName" || key == "isMultiContracotr") return "NA";
                SendErrorEmail("SiteUrl: <br/>" + web.Url + " <br/><br/>" +
                               "Error Message:<br/> Error in GetParameter: No records for Type " + key);
            }
            else
            {
                SendErrorEmail("SiteUrl: <br/>" + web.Url + " <br/><br/>" +
                               "Error Message:<br/> Error in GetParameter: Multiple records for Type " + key);
            }
            return retValue;

        }
        catch (Exception e)
        {
            SendErrorEmail("SiteUrl: <br/>" + web.Url + "<br/><br/>" +
                           "Error Message:<br/> Error in GetParameter Type:" + key + " <br/><br/> " +
                           "Message: <br/>" + e.Message + "<br/><br/>" +
                           "StackTrace: <br/>" + e.StackTrace);
            return "";
        }
        finally
        {
            web.Dispose();
        }
    }

    public static void SendErrorEmail(string Body)
    {
        string from = GetParameter("FromError");
        //string to = GetParameter("AdminEmails");
        List<string> ToUsers = new List<string>();
        List<string> temp = new List<string>();
        temp = GetParameter("ErrorAdmin").Split(',').ToList();
        for (int i = 0; i < temp.Count; i++)
            ToUsers.Add(temp[i].Trim());

        string subject = Environment.MachineName + " - " + GetParameter("ProjectName") + " - Error on PrintView.aspx.cs";
        string body = Body;

        SmtpClient myClient = new SmtpClient(ConfigurationManager.AppSettings["SmtpServer"]);

        if (!string.IsNullOrEmpty(GetParameter("FromErrorSmtpUser")))
            myClient.Credentials = new System.Net.NetworkCredential(GetParameter("FromErrorSmtpUser"), GetParameter("FromErrorSmtpPWD"));
        else
            myClient.UseDefaultCredentials = true;

        MailMessage msg = new MailMessage();
        msg.From = new MailAddress(from, GetParameter("FromErrorName"));
        //msg.To.Add(to);
        for (int i = 0; i < ToUsers.Count; i++) { msg.To.Add(ToUsers[i]); }
        msg.Subject = subject;
        msg.Body = body;
        msg.IsBodyHtml = true;
        myClient.Send(msg);
    }

    public void SetMultiChoiceLookups(SPWeb web, SPListItem Listitem, ref string EmailBody, bool isDesc_Choice)
    {
        SPList list = web.Lists.TryGetList("InspPurpose");
        if (isDesc_Choice)
            list = web.Lists.TryGetList("InspDesc");

        if (list != null)
        {
            SPQuery listQuery = new SPQuery();
            string ListQuery = "";

            #region BUILD THE QUERY
            ListQuery = "<Where><And>";
            if (isDesc_Choice)
                ListQuery += "<And>";

            ListQuery += "<Eq><FieldRef Name='InspType' /><Value Type='LookupMulti'>" + Listitem["InspType"].ToString().Split('#')[1] + "</Value></Eq>" +
                         "<Eq><FieldRef Name='Discipline' /><Value Type='LookupMulti'>" + Listitem["Discipline"].ToString().Split('#')[1] + "</Value></Eq>" +
                        "</And>";


            if (isDesc_Choice)
            {
                SPFieldLookupValueCollection InspPurplkp = (Listitem["InspPurpose"] != null) ? new SPFieldLookupValueCollection(Listitem["InspPurpose"].ToString()) : null;
                if (InspPurplkp != null && InspPurplkp.Count > 0)
                {
                    int CountValues = InspPurplkp.Count;
                    if (CountValues == 1)
                        ListQuery += "<Eq><FieldRef Name='InspPurpose' /><Value Type='LookupMulti'>" + InspPurplkp[0].LookupValue + "</Value></Eq>";
                    else if (CountValues == 2)
                    {
                        ListQuery += "<Or>";
                        for (int i = 0; i < InspPurplkp.Count(); i++)
                        {
                            ListQuery += "<Eq><FieldRef Name='InspPurpose' /><Value Type='Text'>" + InspPurplkp[i].LookupValue + "</Value></Eq>";
                        }
                        ListQuery += "</Or>";
                    }
                    else
                    {
                        ListQuery += CreateNested_OR_Operator(InspPurplkp.Count - 1);
                        int j = 0;
                        for (int i = 0; i < InspPurplkp.Count(); i++)
                        {
                            if (j < 2)
                            {
                                ListQuery += "<Eq><FieldRef Name='InspPurpose' /><Value Type='Text'>" + InspPurplkp[i].LookupValue + "</Value></Eq>";
                                if (j == 1)
                                    ListQuery += "</Or>";
                            }
                            else
                            {
                                ListQuery += "<Eq><FieldRef Name='InspPurpose' /><Value Type='Text'>" + InspPurplkp[i].LookupValue + "</Value></Eq></Or>";
                            }
                            j++;
                        }
                    }
                }
                ListQuery += "</And></Where>";
            }
            else ListQuery += "</Where>";
            ListQuery += "<OrderBy><FieldRef Name='Title' /></OrderBy>";
            #endregion

            listQuery.Query = ListQuery;
            SPListItemCollection listItems = list.GetItems(listQuery);
            if (listItems.Count > 0)
            {
                SPFieldLookupValueCollection LookupColumn = (Listitem["InspPurpose"] != null) ? new SPFieldLookupValueCollection(Listitem["InspPurpose"].ToString()) : null;
                if (isDesc_Choice)
                    LookupColumn = (Listitem["InspDesc"] != null) ? new SPFieldLookupValueCollection(Listitem["InspDesc"].ToString()) : null;
            
                if (LookupColumn != null && LookupColumn.Count > 0)
                {
                    value = "<table style='border-collapse:collapse; color:black;'  width='100%' cellpadding='0' cellspacing='0'>";
                    bool isVisited = false;
                    int colCount = 0;

                    #region SET WIDTH AND NUMBER OF COLUMNS PER ROW BASED ON LENGTH
                    int width = 0, ColPerRow = 0;
                    foreach (SPListItem listItem in listItems)
                    {
                        string Value = listItem["Title"].ToString();
                        if (Value.Length > 1 && Value.Length <= 15)
                        {
                            if (width < 20)
                            {
                                width = 16;
                                ColPerRow = 6;
                            }
                        }
                        else if (Value.Length > 15 && Value.Length <= 30)
                        {
                            if (width < 25)
                            {
                                width = 20;
                                ColPerRow = 5;
                            }
                        }
                        else if (Value.Length > 30 && Value.Length <= 45)
                        {
                            if (width < 33)
                            {
                                width = 25;
                                ColPerRow = 4;
                            }
                        }
                        else
                        {
                            if (width < 50)
                            {
                                width = 33;
                                ColPerRow = 3;
                            }

                        }
                    }
                    #endregion

                    foreach (SPListItem listItem in listItems)
                    {
                        #region SET TICK-MARKS DICTIONARY
                        string Value = listItem["Title"].ToString();
                        if (colCount == ColPerRow)
                        {
                            value += "</tr>";
                            isVisited = false;
                            colCount = 0;
                        }

                        if (!isVisited)
                        {
                            value += "<tr>";
                            isVisited = true;
                        }

                        string CurrentValue = "";
                        bool isFound = false;
                        for (int i = 0; i < LookupColumn.Count(); i++)
                        {
                            CurrentValue = LookupColumn[i].LookupValue;
                            if (Value == CurrentValue)
                            {
                                isFound = true;
                                break;
                            }
                        }

                        if (isFound)
                            value += "<td width = '" + width + "% '><input type='checkbox' name='" + Value + "' value='" + Value + "' checked>" + Value + "</td>";
                        else
                            value += "<td width = '" + width + "%'><input type='checkbox' name='" + Value + "' value='" + Value + "'>" + Value + "</td>";
                        colCount++;
                        #endregion
                    }
                }
            }
            value += "</table>";

        }
        EmailBody = EmailBody.Replace(ColumnInBrackets, value);
    }

    static string CreateNested_OR_Operator(int count)
    {
        string ConcatenatedAND = "";
        int Counter = count + 1;
        for (int i = 1; i < Counter; i++)
        {
            ConcatenatedAND += "<Or>";
        }
        return ConcatenatedAND;
    }

    public byte[] PDFConvert(string html)
    {
        PdfConverter pdfConverter = new PdfConverter();
        pdfConverter.AuthenticationOptions.Username = CredentialCache.DefaultNetworkCredentials.UserName;
        pdfConverter.AuthenticationOptions.Password = CredentialCache.DefaultNetworkCredentials.Password;
        pdfConverter.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4;
        pdfConverter.PdfDocumentOptions.PdfPageOrientation = PDFPageOrientation.Portrait;

        //pdfConverter.PdfDocumentOptions.ShowHeader = true;
        //pdfConverter.PdfHeaderOptions.DrawHeaderLine = false;
        //pdfConverter.PdfDocumentOptions.ShowFooter = true;
        //pdfConverter.PdfDocumentOptions.LeftMargin = 15;
        //pdfConverter.PdfDocumentOptions.RightMargin = 15;
        //pdfConverter.PdfDocumentOptions.TopMargin = -15;//20;
        //pdfConverter.PdfDocumentOptions.BottomMargin = 10;// 21;

        pdfConverter.PdfDocumentOptions.PdfCompressionLevel = PdfCompressionLevel.NoCompression;//PdfCompressionLevel.Normal;//PdfCompressionLevel.Best; //


        /*
        pdfConverter.PdfDocumentOptions.FitWidth = true;
        pdfConverter.PdfDocumentOptions.EmbedFonts = true;

        pdfConverter.PdfDocumentOptions.GenerateSelectablePdf = true;
        pdfConverter.PdfDocumentOptions.JpegCompressionEnabled = true;
        
        pdfConverter.OptimizePdfPageBreaks = true;
        pdfConverter.AvoidTextBreak = true;
        pdfConverter.NavigationTimeout = 3600;
        pdfConverter.ConversionDelay = 3;
        pdfConverter.JavaScriptEnabled = true;
        */

        pdfConverter.LicenseKey = "NB8GFAYHFAQUAhoEFAcFGgUGGg0NDQ0=";

        byte[] ret = pdfConverter.GetPdfBytesFromHtmlString(html);
        return ret;
    }

    public static string GetVersionedMultiLineTextAsPlainText(SPListItem item, string key, ref int Count, bool isLatest, bool Asc)
    {
        StringBuilder sb = new StringBuilder();
        string substr = "";
        ArrayList arrlist = new ArrayList();
        foreach (SPListItemVersion version in item.Web.Lists[item.ParentList.ID].Items[item.UniqueId].Versions)
        {
            substr = "";
            SPFieldMultiLineText field = version.Fields[key] as SPFieldMultiLineText;
            if (field != null)
            {
                string comment = field.GetFieldValueAsText(version[key]);
                if (comment != null && comment.Trim() != string.Empty)
                {
                    if (!isLatest)
                    {
                        substr += "" + version.CreatedBy.User.Name + " (" + version.Created.ToString("MM/dd/yyyy hh:mm tt") + ") " + comment;
                        substr += "<br/><br/>";
                    }
                    else substr = comment;


                    if (isLatest)
                    {
                        sb.Append(substr);
                        break;
                    }
                    else
                    {

                        if (Asc)
                        {
                            arrlist.Add(substr);
                        }
                        else
                        {
                            sb.Append(substr);
                            Count++;
                        }
                    }
                }
            }
        }

        if(Asc && arrlist != null && arrlist.Count > 0)
        {
            sb = new StringBuilder();
            for (int i = arrlist.Count -1; i >= 0; i--)
            {
                if (arrlist[i] != null && arrlist[i].ToString() != "")
                    sb.Append(arrlist[i].ToString());
            }
        }

        return sb.ToString();
    }

}