using System;
using System.Data;
using System.Configuration;
using System.Collections;
using System.Web;
using System.Web.UI.WebControls;
using Microsoft.SharePoint;
using System.IO;

using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.security;

using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Crypto;

using Microsoft.SharePoint.Utilities;
using System.Collections.Generic;
using System.Net.Mail;
using iTextSharpSign;

using Winnovative.WnvHtmlConvert;
using System.Net;
using System.Text.RegularExpressions;
using System.Linq;
using Microsoft.SharePoint.Administration;

public partial class RIWCompile : System.Web.UI.Page
{
    #region Global variables

    public static string ErrorSubject = " - PROJ - EInspection - Error on RIWCompile.aspx.cs";
    public static string SmtpServer = ConfigurationManager.AppSettings["SmtpServer"];
    public static string layoutPath = "_layouts/15/PCW/General/EForms";
    public static string LeadRedirectURL = "/Lists/Inspection Lead Tasks/";
    public string redirectURL = "";

    public string Referrer = "";
    string RIWQueryStrID = "", LeadQueryStrID = "";
    string MainList = "", LeadTaskList = "", PartTaskList = "";
    string StrFullRef = "", ARE = "";
    //string Pdftemplatepath = ConfigurationSettings.AppSettings["PdfFilePath"].ToString();
    SPList objlist = null;
    SPList objlistLeadTasks = null;
    SPList objWorkflowTask = null;
    SPListItem RIWItem = null;
    SPListItem LeadItem = null;
    bool ApplyDigitalSign = true;
    bool BlankCommentsForm = false;

    public static string LeadName = "";
    public string Discipline = "";
    public DateTime DateClosed ;
    public string UserName = "";

    #endregion

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

    public static SPWeb objweb = null;
    public static int Version = GetSPVersion();

    protected void Page_Load(object sender, EventArgs e)
    {
        try
        {
            #region SESSION EXPIRED THEN REDIRECT TO LOGIN
            if (SPContext.Current.Web.CurrentUser == null)
            {
                string serverRelativeUrl = SPContext.Current.Web.ServerRelativeUrl;
                string layoutUrl = HttpContext.Current.Request.Url.PathAndQuery;
                //string CurrentUrl = System.Web.HttpUtility.UrlEncode(serverRelativeUrl + layoutUrl);
                string CurrentUrl = serverRelativeUrl + layoutUrl;
                //Response.Write(layoutUrl);
                //string LoginURL = "/_login/default.aspx?ReturnUrl=" + SPContext.Current.Web.Url.Replace("/", "%2f") + "%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d" + CurrentUrl.Replace("/", "%252D") + "&Source=" + CurrentUrl.Replace("/", "%2F");
                string LoginURL = "/_login/default.aspx?ReturnUrl=" + serverRelativeUrl + "%2f_layouts%2fAuthenticate.aspx%3fSource%3d" + CurrentUrl + "&Source=" + CurrentUrl;

                if (Version > 14)
                    LoginURL = "/_login/default.aspx?ReturnUrl=" + serverRelativeUrl + "%2f_layouts%2f15%2fAuthenticate.aspx%3fSource%3d" + CurrentUrl + "&Source=" + CurrentUrl;

                Response.Write("<script>window.location = '" + LoginURL + "';</script>");
                return;
            }
            #endregion

            DateTime Start = DateTime.Now;
            SPFieldUserValue LeadUser = null;
            string TimeLog = "";

            if (!IsPostBack)
            {
                if (Request.UrlReferrer != null)
                    Referrer = Request.UrlReferrer.ToString();
            }
            MainList = "Inspection Request";
            LeadTaskList = "Inspection Lead Tasks";
            PartTaskList = "Inspection Part Tasks";
            LeadName = SPContext.Current.Web.CurrentUser.Name;
            string UserEmail = "";
            try { UserEmail = SPContext.Current.Web.CurrentUser.Email; }
            catch { }

            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    objweb = site.OpenWeb();
                }

                objweb.AllowUnsafeUpdates = true;
            //});

            //ProjectDesc.Text = "E-Inspection Review by Lead Action";

            //if (SPContext.Current != null && SPContext.Current.Web != null && SPContext.Current.Web.CurrentUser != null)
            //    this.UserName = SPContext.Current.Web.CurrentUser.Name;

            #region QUERY E-Inspection AND LEAD
            objlist = objweb.Lists[MainList];
            objlistLeadTasks = objweb.Lists[LeadTaskList];
            objWorkflowTask = objweb.Lists[PartTaskList];

            loader.ImageUrl = objweb.Url + ConfigurationManager.AppSettings["loaderUrl"];


            //string RIWViewFields =//"<FieldRef Name='ID' />" +
            //                        " <FieldRef Name='Attachments' />" +
            //                        "<FieldRef Name='Title' />" +
            //                        "<FieldRef Name='date_closed' />" +
            //                        "<FieldRef Name='FullRef' />" +
            //                        "<FieldRef Name='project_name' />" +
            //                        "<FieldRef Name='project_name_x003a_Title' />" +
            //                        "<FieldRef Name='discipline' />" +
            //                        "<FieldRef Name='LeadAction' />" +
            //                        "<FieldRef Name='bldgar_nbr' />" +
            //                        "<FieldRef Name='sub_name' />" +
            //                        "<FieldRef Name='attach_dwgskt1' />" +
            //                        "<FieldRef Name='ItemOfWork' />" +
            //                        "<FieldRef Name='Quantity' />" +
            //                        "<FieldRef Name='Unit' />";

            RIWQueryStrID = HttpContext.Current.Request.QueryString["RIWID"];
                LeadQueryStrID = HttpContext.Current.Request.QueryString["ID"];

                int LeadItemID = 0;
                int RIWItemID = 0;
                if (!String.IsNullOrEmpty(RIWQueryStrID))
                {
                    DateTime Start1 = DateTime.Now;
                    if (int.TryParse(RIWQueryStrID, out RIWItemID))
                    {
                        RIWItemID = int.Parse(RIWQueryStrID);
                        // RIWItem = objlist.GetItemById(RIWItemID);

                        SPQuery _query = new SPQuery();
                        _query.Query = "<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + RIWItemID + "</Value></Eq></Where>";
                        //_query.ViewFields = RIWViewFields;
                        _query.ViewAttributes = "Scope=\"Recursive\"";
                        SPListItemCollection RIWItems = objlist.GetItems(_query);
                        if (RIWItems.Count == 1)
                            RIWItem = RIWItems[0];
                        else
                        {
                            MessageAndRedirect("This RIW has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);
                            return;
                        }

                        StrFullRef = (RIWItem["FullRef"] != null) ? RIWItem["FullRef"].ToString() : "";
                        Discipline = (RIWItem["Discipline"] != null) ? RIWItem["Discipline"].ToString().Split('#')[1] : "";
                        redirectURL = LeadRedirectURL;
                }
                    TimeLog += "GET ITEM FROM MAIN RIW - Page_Load - " + DateTime.Now.Subtract(Start1).TotalSeconds + "<br/><br/>";
                }
                else if (!String.IsNullOrEmpty(LeadQueryStrID))
                {
                    DateTime Start2 = DateTime.Now;
                    if (int.TryParse(LeadQueryStrID, out LeadItemID))
                    {
                        {
                            LeadItemID = int.Parse(LeadQueryStrID);
                            LeadItem = objlistLeadTasks.GetItemById(LeadItemID);
                            StrFullRef = (LeadItem["Reference"] != null) ? LeadItem["Reference"].ToString() : "";
                            Discipline = (LeadItem["Trade"] != null) ? LeadItem["Trade"].ToString() : "";
                            redirectURL = LeadRedirectURL;//"/Lists/RIW/";
                            LeadUser = (LeadItem["AssignedTo"] != null) ? new SPFieldUserValue(objweb, LeadItem["AssignedTo"].ToString()) : null;
                            //LeadName = LeadUser.User.Name;// SPContext.Current.Web.CurrentUser.Name;//Used in getting the signature image file

                    }
                    }
                    else //contains text means fullRef
                        StrFullRef = HttpContext.Current.Request.QueryString["ID"];
                    TimeLog += "GET ITEM FROM LEAD RIW - Page_Load - " + DateTime.Now.Subtract(Start2).TotalSeconds + "<br/><br/>";
                }

                //Get RIW and Lead Item by Ref
                if (RIWItem == null)
                {
                    DateTime Start3 = DateTime.Now;
                    SPQuery _query = new SPQuery();
                    _query.Query = "<Where><Eq><FieldRef Name='FullRef' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
                    _query.ViewAttributes = "Scope=\"Recursive\"";
                    //_query.ViewFields = RIWViewFields;
                    SPListItemCollection RIWItems = objlist.GetItems(_query);
                    if (RIWItems.Count == 1)
                        RIWItem = RIWItems[0];
                    else
                    {
                        MessageAndRedirect("This RIW has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);
                        return;
                    }
                    TimeLog += "GET ITEM FROM MAIN RIW IF NULL - Page_Load - " + DateTime.Now.Subtract(Start3).TotalSeconds + "<br/><br/>";
                }
                if (LeadItem == null)
                {
                    DateTime Start4 = DateTime.Now;
                    SPQuery _Leadquery = new SPQuery();
                    _Leadquery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
                    _Leadquery.ViewAttributes = "Scope=\"Recursive\"";
                    SPListItemCollection LeadItems = objlistLeadTasks.GetItems(_Leadquery);
                    if (LeadItems.Count == 1)
                        LeadItem = LeadItems[0];
                    else
                    {
                        MessageAndRedirect("This RIW has no lead task. Please contact PWS administrator.", SPContext.Current.Web.Url + redirectURL);
                        return;
                    }
                    TimeLog += "GET ITEM FROM Lead RIW IF NULL - Page_Load - " + DateTime.Now.Subtract(Start4).TotalSeconds + "<br/><br/>";
                }
            #endregion

            #region IS ALLOWED TO CLOSE THE TASK
            DateClosed = (RIWItem["SentToContractorDate"] != null) ? (DateTime)RIWItem["SentToContractorDate"] : DateTime.Now;
            SPFieldUserValueCollection AssignedToLead = (LeadItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(objweb, LeadItem["AssignedTo"].ToString()) : null;
            bool goAhead = false;
            for (int i = 0; i < AssignedToLead.Count; i++)
            {
                if (AssignedToLead[i].User.Email.ToString().ToLower() == UserEmail.ToLower())
                    goAhead = true;
            }

            if (!goAhead)
                MessageAndRedirect("This task is not assigned to you.", SPContext.Current.Web.Url + redirectURL);
            #endregion

            if (!IsPostBack)
            {
                DateTime Start6 = DateTime.Now;

                DateTime Start6_1 = DateTime.Now;
                #region Creating Data Table with required columns for Final Output.

                DataTable dtstore = new DataTable();
                dtstore.Columns.Add("ID", typeof(string));
                dtstore.Columns.Add("Title", typeof(string));
                dtstore.Columns.Add("AssignedTo", typeof(string));
                //dtstore.Columns.Add("SeniorEngineer", typeof(string));
                //dtstore.Columns.Add("PartRE", typeof(string));
                dtstore.Columns.Add("Date", typeof(DateTime));
                dtstore.Columns.Add("Status", typeof(string));
                dtstore.Columns.Add("Code", typeof(string));
                //dtstore.Columns.Add("CMQuantity", typeof(string));
                //dtstore.Columns.Add("CMUnit", typeof(string)); 
                dtstore.Columns.Add("Comment", typeof(string));
                //dtstore.Columns.Add("DaysHeld", typeof(string));
                //dtstore.Columns.Add("WorkflowLink", typeof(string));

                DataTable dtFinal = new DataTable();
                dtFinal.Columns.Add("ID", typeof(string));
                dtFinal.Columns.Add("Title", typeof(string));
                dtFinal.Columns.Add("AssignedTo", typeof(string));
                //dtFinal.Columns.Add("SeniorEngineer", typeof(string));
                //dtFinal.Columns.Add("PartRE", typeof(string));
                dtFinal.Columns.Add("Date", typeof(DateTime));
                dtFinal.Columns.Add("Status", typeof(string));
                dtFinal.Columns.Add("Code", typeof(string));
                //dtFinal.Columns.Add("CMQuantity", typeof(string));
                //dtFinal.Columns.Add("CMUnit", typeof(string));
                dtFinal.Columns.Add("Comment", typeof(string));
                //dtFinal.Columns.Add("DaysHeld", typeof(string));
                //dtFinal.Columns.Add("WorkflowLink", typeof(string));

                #endregion
                TimeLog += "checking the Status - Creating Data Table - Page_Load - " + DateTime.Now.Subtract(Start6_1).TotalSeconds + "<br/><br/>";

                #region checking the status of the IR from RIW list.

                string code = (LeadItem["Code"] != null) ? LeadItem["Code"].ToString() : "";

                if (code != "")
                    btnUpdateAREMergePdf.Visible = true;

                if (LeadItem.Attachments.Count == 1)
                    SubmitCloseIR.Visible = true;


                if (LeadItem["Status"].ToString() == "Completed")
                {
                    MessageAndRedirect("This Inspection has already been closed.", SPContext.Current.Web.Url + redirectURL);
                    //Response.Write("<script>alert('This RIW has already been closed.');</script>");
                    //Response.Write("<script>window.returnValue = true;window.close();</script>");
                    return;
                }
                //DataTable dt = RIWItems.GetDataTable();
                DateTime Start6_2 = DateTime.Now;
                #region Populate the Inspection info in grid

                lblIRNo.Text = (RIWItem["FullRef"] != null) ? RIWItem["FullRef"].ToString() : "";
                //lblProjectNo.Text = (RIWItem["project_name_x003a_Title"] != null) ? RIWItem["project_name_x003a_Title"].ToString().Split('#')[1] : "";
                lblTrade.Text = (RIWItem[("Discipline")] != null) ? RIWItem[("Discipline")].ToString().Split('#')[1] : "";
                lblRE.Text = LeadName; //(RIWItem["RE_Name"] != null) ? RIWItem["RE_Name"].ToString() : "";
                //lblFacility.Text = (RIWItem["bldgar_nbr"] != null) ? RIWItem["bldgar_nbr"].ToString().Split('#')[1] : "";
                //lblContractor.Text = (RIWItem["sub_name"] != null) ? RIWItem["sub_name"].ToString() : "";
                //DrawingRefVal.Text = (RIWItem["attach_dwgskt1"] != null) ? RIWItem["attach_dwgskt1"].ToString() : "";

                //lblItemOfWork.Text = (RIWItem["ItemOfWork"] != null) ? RIWItem["ItemOfWork"].ToString() : "";
                //lblQuantity.Text = (RIWItem["Quantity"] != null) ? RIWItem["Quantity"].ToString() : "";
                //lblUnit.Text = (RIWItem["Unit"] != null) ? RIWItem["Unit"].ToString() : "";

                lblRIWTitle.Text = (LeadItem["Title"] != null) ? LeadItem["Title"].ToString() : "";

                ddlcode.SelectedIndex = ddlcode.Items.IndexOf(ddlcode.Items.FindByText(code));
                PartComments.Text = (LeadItem["PartComments"] != null) ? SPHttpUtility.ConvertSimpleHtmlToText(LeadItem["PartComments"].ToString(), -1) : "";
                CommentBox.Text = (LeadItem["Comment"] != null) ? LeadItem["Comment"].ToString() : "";

                //CHECK APPROVED DRAWINGS
                //string Drawing_Matched = (LeadItem["DrawingMatched"] != null) ? LeadItem["DrawingMatched"].ToString() : "";
                //ApprovedDwgbtn.SelectedIndex = ApprovedDwgbtn.Items.IndexOf(ApprovedDwgbtn.Items.FindByText(Drawing_Matched));

                //string CMUnit = (LeadItem["CMUnit"] != null) ? LeadItem["CMUnit"].ToString() : "";
                //UnitDDL.SelectedIndex = UnitDDL.Items.IndexOf(UnitDDL.Items.FindByText(CMUnit));

                //string CMQuantity = (LeadItem["CMQuantity"] != null) ? LeadItem["CMQuantity"].ToString() : "";
                //QtyTextBox.Text = CMQuantity;

                if (LeadItem.Attachments.Count > 0)
                {
                    HyperLinkLeadAttach.NavigateUrl = LeadItem.Attachments.UrlPrefix + LeadItem.Attachments[0];
                    HyperLinkLeadAttach.Text = LeadItem.Attachments[0];
                    lblmsg.Text = "Warning this Lead Task has attachment and pressing Merge again may overwrite the existing attachement.";
                }
                //lblARE.Text = dt.Rows[0]["SeniorEng"].ToString();  

                //SPFieldUrlValue value = new SPFieldUrlValue(RIWItem["PDFLink"].ToString().Trim());
                if (RIWItem.Attachments.Count == 0)
                {
                    //MessageAndRedirect("This RIW has no attachment to get the cover page from, please close it from Lead Task.", SPContext.Current.Web.Url + LeadRedirectURL + "DispForm.aspx?ID=" + LeadItemID.ToString());
                    Response.Write("<script>alert('This E-Inspection has no attachment to get the cover page from, please close it from Lead Task.');</script>");
                    Response.Write("<script>window.location = '" + SPContext.Current.Web.Url + LeadRedirectURL + "DispForm.aspx?ID=" + LeadItem.ID.ToString() + "';</script>");
                    return;
                }
                else if (RIWItem.Attachments.Count == 1)
                {
                    if (RIWItem.Attachments[0].ToLower().EndsWith("pdf"))
                    {
                        Pdfhyplink.NavigateUrl = RIWItem.Attachments.UrlPrefix + RIWItem.Attachments[0];
                        Pdfhyplink.Text = RIWItem.Attachments[0];
                    }
                    else
                    {
                        lblmsg.Text = "The contractor attachment is not of type PDF, you must close this Inspection from lead task.";
                        return;
                    }
                    //Pdfhyplink.NavigateUrl = (value.Url.Contains(" ") ? value.Url.Replace(" ", "") : value.Url);
                    //Pdfhyplink.Text = "PdfLink";
                }
                else
                {
                    lblmsg.Text = "This Inspection has multiple attachments, you must close this RIW from lead task.";
                    return;
                }
                #endregion
                TimeLog += "checking the Status - Populate the Inspection info in grid - Page_Load - " + DateTime.Now.Subtract(Start6_2).TotalSeconds + "<br/><br/>";

                DateTime Start6_3 = DateTime.Now;
                #region getting the records from Part Tasks

                SPQuery _WTQuery = new SPQuery();
                _WTQuery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
                _WTQuery.ViewAttributes = "Scope=\"Recursive\"";
                //_WTQuery.ViewFields = "<FieldRef Name='ID' />" +
                //                    "<FieldRef Name='Title' />" +
                //                    "<FieldRef Name='AssignedTo' />" +
                //                    "<FieldRef Name='PartRE' />" +
                //                    "<FieldRef Name='Date' />" +
                //                    "<FieldRef Name='Status' />" +
                //                    "<FieldRef Name='Code' />" +
                //                    "<FieldRef Name='CMQuantity' />" +
                //                    "<FieldRef Name='CMUnit' />" +
                //                    "<FieldRef Name='Comment' />";

                DataTable dtTask = objWorkflowTask.GetItems(_WTQuery).GetDataTable();
                SPListItemCollection PartTasks = objWorkflowTask.GetItems(_WTQuery);

                if (dtTask != null)
                {
                    //Response.Write("<script>alert('This IR has part " + dtTask.Rows.Count.ToString() + ".');</script>");
                    string strQueryFilter = "";
                    if (dtTask.Rows.Count > 0)
                    {
                        #region filling the dtstore table.
                        for (int i = 0; i < dtTask.Rows.Count; i++)
                        {
                            DataRow dr = dtstore.NewRow();
                            dr["ID"] = dtTask.Rows[i]["ID"].ToString();
                            dr["Title"] = dtTask.Rows[i]["Title"].ToString();
                            dr["AssignedTo"] = dtTask.Rows[i]["AssignedTo"].ToString();
                            //Get User collection
                            if (dtTask.Rows[i]["AssignedTo"].ToString().Contains(";#"))
                            {
                                SPListItem taskItem = objWorkflowTask.GetItemById(int.Parse(dtTask.Rows[i]["ID"].ToString()));
                                SPFieldUserValueCollection usrColl = (taskItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(objweb, taskItem["AssignedTo"].ToString()) : null;
                                string PartNames = "";
                                foreach (SPFieldUserValue usrval in usrColl)
                                    if (!PartNames.Contains(usrval.User.Name)) PartNames += usrval.User.Name + "<br/>";
                                //if (PartNames != "") PartNames += "<br/> " + usrval.User.Name ;
                                //else PartNames += usrval.User.Name;
                                dr["AssignedTo"] = PartNames;
                            }

                            //dr["SeniorEngineer"] = dtTask.Rows[i]["SeniorEngineer"].ToString();
                            //dr["PartRE"] = dtTask.Rows[i]["PartRE"].ToString();
                            dr["Date"] = dtTask.Rows[i]["Date"];
                            dr["Status"] = dtTask.Rows[i]["Status"].ToString();
                            dr["Code"] = dtTask.Rows[i]["Code"].ToString();
                            //dr["CMQuantity"] = dtTask.Rows[i]["CMQuantity"].ToString();
                            //dr["CMUnit"] = dtTask.Rows[i]["CMUnit"].ToString();
                            dr["Comment"] = dtTask.Rows[i]["Comment"].ToString();
                            //dr["WorkflowLink"] = Regex.Match(dtTask.Rows[i]["WorkflowLink"].ToString(), @"ID=(\d+)").Groups[1].Value;

                            //if (dtTask.Rows[i]["Status"].ToString() == "Completed")
                            //    dr["DaysHeld"] = "";
                            //else
                            //{
                            //    DateTime futurDate = Convert.ToDateTime(dtTask.Rows[i]["Date"]);
                            //    DateTime TodayDate = DateTime.Now;
                            //    double numberOfDays = (TodayDate - futurDate).TotalDays;
                            //    int number = Convert.ToInt32(numberOfDays);
                            //    dr["DaysHeld"] = number.ToString();
                            //}

                            dtstore.Rows.Add(dr);
                            //ViewState["AssignedTo"] = dtTask.Rows[i]["AssignedTo"].ToString();
                        }
                        #endregion

                        DataRow[] result = dtstore.Select(strQueryFilter);
                        if (result.Length > 0)
                        {
                            foreach (DataRow row in result)
                            {
                                dtFinal.Rows.Add(row.ItemArray);
                            }
                            ViewState["FinalData"] = dtFinal;
                            grdWorkflowtasks.DataSource = dtFinal;
                            grdWorkflowtasks.DataBind();
                        }
                    }
                }
                #endregion
                TimeLog += "checking the Status - getting the records from Part Tasks - Page_Load - " + DateTime.Now.Subtract(Start6_3).TotalSeconds + "<br/><br/>";

                //btnUpdateAREMergePdf.Visible = true;
                //btnUpdateAREMergePdfNoWT.Visible = true;
                //btnSign.Visible = true;
                //PanelRE.Visible = true;
                //btnCloseIR.Visible = true;
                #endregion

                TimeLog += "Total checking the status of the IR from RIW list. - Page_Load - " + DateTime.Now.Subtract(Start6).TotalSeconds + "<br/><br/>";
            }

            int total = (int)DateTime.Now.Subtract(Start).TotalSeconds;
            if (total > 5)
            {
                TimeLog += "TOTAL TIME = " + total;

                //SendErrorEmail(StrFullRef + " - Page_Load<br/><br/>" + TimeLog, "RIW Lead Page Load");
            }
        }
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail("Page Load Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> RIW Ref:" + lblIRNo.Text + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        //finally { if (objweb != null) objweb.Dispose(); }
    }    

    protected void btnUpdateAREMergePdf_Click(object sender, EventArgs e)
    {
        SPWeb objweb = null;
        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Site.ID))
                { objweb = site.OpenWeb(); }
            });

            objweb.AllowUnsafeUpdates = true;

            lblmsg.Text = "";
            string UserName = "";//pdfPath = "",strpathSplit1 = "",strpathSplit2 = "", strpath3 = "",
            int intotherfilescount = 0, intEngineersComments = 0, Satuscount = 0, WTask = 0, intimagecount = 0; //IRPagescount = 0,
            //SPSecurity.RunWithElevatedPrivileges(delegate()            {
            //    using (SPWeb objweb = SPContext.Current.Site.OpenWeb())                {

            SPList objlist = objweb.Lists[PartTaskList];
            //SPDocumentLibrary _objInspecRepository = objweb.Lists.TryGetList("Inspection Repository") as SPDocumentLibrary;

            #region Checking the IR and Attachments count
            //if (string.IsNullOrEmpty(UnitDDL.SelectedValue))
            //{
            //    lblmsg.Text = "Not saved, Please select Unit first.";
            //    return;
            //}

            //if (string.IsNullOrEmpty(QtyTextBox.Text.Trim()))
            //{
            //    lblmsg.Text = "Not saved, Please select Quantity first.";
            //    return;
            //}

            if (ddlcode.SelectedValue == "Select Code")
            {
                lblmsg.Text = "Please select code first.";
                return;
            }
            if (CommentBox.Text == "")
            {
                lblmsg.Text = "Please enter final comments.";
                return;
            }
            foreach (GridViewRow gritemcheck in grdWorkflowtasks.Rows)
            {
                CheckBox CheckBoxWTask = (CheckBox)gritemcheck.FindControl("CheckBoxWTask");
                Label lblIDcheck = (Label)gritemcheck.FindControl("lblID");
                Label lblWTStatus = (Label)gritemcheck.FindControl("lblWTStatus");
                Label lblAssignedTo = (Label)gritemcheck.FindControl("lblAssignedTo");
                //Label lblSeniorEngineer = (Label)gritemcheck.FindControl("lblSeniorEngineer");
                CheckBoxList chkBxListcheck = (CheckBoxList)gritemcheck.FindControl("Chkattachments");

                if (CheckBoxWTask.Checked)
                {
                    WTask++;
                    if (lblWTStatus.Text != "Completed")
                    {
                        #region Checking whether Inspector completed his task or not.

                        if (UserName == "")
                            UserName += lblAssignedTo.Text + " ";
                        else
                            UserName += " , " + lblAssignedTo.Text + " ";

                        Satuscount++;

                        #endregion
                    }
                    else if (lblWTStatus.Text == "Completed")
                    {
                        #region Checking whether Engineer completed his task or not.
                        if (chkBxListcheck != null)
                        {
                            if (chkBxListcheck.Items.Count == 0)
                            {
                                SPListItem item = objlist.GetItemById(Convert.ToInt32(lblIDcheck.Text));
                                SPAttachmentCollection _attachments = item.Attachments;
                                if (_attachments != null)
                                {
                                    if (_attachments.Count != 0)
                                    {
                                        if (UserName == "")
                                            UserName += lblAssignedTo.Text + " ";
                                        else
                                            UserName += " , " + lblAssignedTo.Text + " ";

                                        Satuscount++;
                                    }
                                }
                            }

                        }
                        #endregion
                    }
                    if (chkBxListcheck != null)
                    {
                        foreach (System.Web.UI.WebControls.ListItem itemcheck in chkBxListcheck.Items)
                        {
                            if (itemcheck.Selected)
                            {
                                string UpperCaseFilename = itemcheck.Value.ToUpper();
                                intotherfilescount++;

                                //if (itemcheck.Value.Contains(lblIRNo.Text) && itemcheck.Value.ToUpper().Contains(".PDF"))
                                //{
                                //    intonlyIRPdfcheck++;                                               
                                //}
                                //else 
                                if (UpperCaseFilename.Contains(".PDF")) // (itemcheck.Value.ToLower().Contains("engineer")) &&
                                {
                                    intEngineersComments++;
                                }
                                else if (UpperCaseFilename.Contains(".JPG") || UpperCaseFilename.Contains(".PNG")) // (itemcheck.Value.ToLower().Contains("engineer")) &&
                                {
                                    intimagecount++;
                                }
                            }
                        }
                    }
                }
            }
            #endregion

            #region Validation of counts
            //string PDFpagesToInclude = rblist_PDFpagesToInclude.SelectedValue;
            //if (WTask == 0)
            //{
            //    if (PDFpagesToInclude == "CoverPageFromRIW")
            //    {
            //        lblmsg.Text = "You have chosen to include only cover page from contractor RIW, Please select at least one part task which has PDF attachments in order to close the RIW.";
            //        return;
            //    }
            //}

            if (Satuscount > 0)
            {
                lblmsg.Text = "Selected Part Task Pending from " + UserName + ". To ignore and close anyway, please unselect the open tasks.";
                return;
            }

            //if (intEngineersComments == 0)
            //{
            //    if (PDFpagesToInclude == "CoverPageFromRIW")
            //    {
            //        lblmsg.Text = "You have chosen to include only cover page from contractor RIW, you should select a part task PDF file.";// (OR) The file naming is wrong.";
            //        return;
            //    }
            //}
            #endregion

            #region COMMENTED Code for Updating the IR Pdf without attachment
            //if (intotherfilescount == 0)
            //{
            //SPDocumentLibrary DocRepositoryonlypdf = objweb.Lists.TryGetList("Repository") as SPDocumentLibrary;
            ////Creating Subfolder inside folder
            //string subFolderUrlonlypdf = SPContext.Current.Web.Url + "/Repository/ReviewedDeliverables/IR";

            //SPList _listSLODonlypdf = objweb.Lists[MainList];

            //string strCDSNumber = "";
            //SPListItem _itemSLOD = _listSLODonlypdf.GetItemById(Convert.ToInt32(HttpContext.Current.Request.QueryString["ID"]));
            //if (_itemSLOD != null)
            //{
            //    strCDSNumber = _itemSLOD["CDSNumber"].ToString();
            //    var linkItem = new SPFieldUrlValue(_itemSLOD["PDFLink"].ToString());
            //    pdfPath = linkItem.Url;
            //}
            //SPListItem subFolder = null;
            //string subsubFolderUrl = "";
            //if (SPContext.Current.Web.GetFolder(subFolderUrlonlypdf + "/" + strCDSNumber).Exists)
            //{
            //    subsubFolderUrl = subFolderUrlonlypdf + "/" + strCDSNumber;
            //}
            //else
            //{
            //    subFolder = DocRepositoryonlypdf.Items.Add(subFolderUrlonlypdf, SPFileSystemObjectType.Folder, strCDSNumber);
            //    subFolder.Update();
            //    subsubFolderUrl = subFolderUrlonlypdf + "/" + strCDSNumber;
            //}


            //string subsubsubFolderUrl = "";
            //if (SPContext.Current.Web.GetFolder(subFolderUrlonlypdf + "/" + strCDSNumber + "/" + lblTrade.Text).Exists)
            //{
            //    subsubsubFolderUrl = subFolderUrlonlypdf + "/" + strCDSNumber + "/" + lblTrade.Text;
            //}
            //else
            //{
            //    SPListItem subsubFolder = DocRepositoryonlypdf.Items.Add(subsubFolderUrl, SPFileSystemObjectType.Folder, lblTrade.Text);
            //    subsubFolder.Update();
            //    subsubsubFolderUrl = subFolderUrlonlypdf + "/" + strCDSNumber + "/" + lblTrade.Text;
            //}

            //SPFile _itemAttachmentonlypdf = objweb.GetFile(pdfPath);
            //using (Stream filestream = _itemAttachmentonlypdf.OpenBinaryStream())
            //{
            //    pdflink.Value = subsubsubFolderUrl + "/" + lblIRNo.Text + ".pdf";
            //    SPFile _myfinalIRonlypdf = DocRepositoryonlypdf.RootFolder.Files.Add(subsubsubFolderUrl + "/" + lblIRNo.Text + ".pdf", filestream, true);
            //    _myfinalIRonlypdf.Update();
            //    objweb.AllowUnsafeUpdates = false;
            //    lblmsg.Text = "";
            //    hyplink.NavigateUrl = pdflink.Value.Split(',')[0];
            //    hyplink.Text = "Download IR Document";
            //    lblmsg.Text = "";
            //    //btnUpdatePdf.Enabled = false;
            //    btnUpdateAREMergePdf.Visible = false;
            //    btnUpdateAREMergePdfNoWT.Visible = false;
            //    btnSign.Visible = true;
            //    PanelRE.Visible = true;
            //    btnCloseIR.Visible = false;
            //    //btnUpdatePdf.Visible = false;
            //    //btnSign.Visible = true;
            //    //hyplink.NavigateUrl = pdflink.Value.Split(',')[0];
            //    //hyplink.Text = "Download IR Document";

            //    //byte[] bytes = _itemAttachmentonlypdf.OpenBinary();

            //    //Response.Clear();
            //    //Response.ContentType = "application/pdf";
            //    //Response.OutputStream.Write(bytes, 0, bytes.Length);
            //    //Response.Flush();
            //    //Response.End();
            //}
            //}
            //else if (intonlyIRPdfcheck == 1 && intEngineersComments > 0)
            //else
            //if (intEngineersComments > 0) {
            #endregion

            #region Code for Updating Lead Task with Pdf attachment

            Label2.Text = "";
            List<String> ListPdffiles = new System.Collections.Generic.List<String>();
            ListPdffiles.Add(RIWItem.Attachments.UrlPrefix + RIWItem.Attachments[0]);

            //if (PDFpagesToInclude == "AllPagesFromRIW")
            //{
            //    ListPdffiles.Add(RIWItem.Attachments.UrlPrefix + RIWItem.Attachments[0]);
            //}

            #region Bind Images and get commented pdf attachments

            //string strpathImages1 = string.Format(Pdftemplatepath + "\\tempPDFImg1_{0}.pdf", _postFix);
            //string _postFix = DateTime.Now.ToString("yyyyMMddHHmmssffff");
            //Create target folder
            //if (!System.IO.Directory.Exists(Pdftemplatepath)) System.IO.Directory.CreateDirectory(Pdftemplatepath);

            Document doc = new Document(PageSize.A4, 0, 0, 70, 60);
            MemoryStream BindImagesFileStream = null;

            if (intimagecount > 0)
            {
                //FileStream _filestream = new FileStream(strpathImages1, FileMode.Create);
                BindImagesFileStream = new MemoryStream();
                PdfWriter _pdfwriter = PdfWriter.GetInstance(doc, BindImagesFileStream);
                doc.Open();
            }
            foreach (GridViewRow gritem in grdWorkflowtasks.Rows)
            {
                CheckBox CheckBoxWTask = (CheckBox)gritem.FindControl("CheckBoxWTask");
                CheckBoxList chkBxList = (CheckBoxList)gritem.FindControl("Chkattachments");
                //Label lblID = (Label)gritem.FindControl("lblID");
                //Label lblAttachmentsLink = (Label)gritem.FindControl("lblAttachmentsLink");
                if (CheckBoxWTask.Checked)
                {
                    if (chkBxList != null)
                    {
                        foreach (System.Web.UI.WebControls.ListItem item in chkBxList.Items)
                        {

                            if (item.Selected)
                            {
                                //if (item.Value.ToLower().Contains("engineer") && item.Value.ToUpper().Contains(".PDF"))
                                if (item.Value.ToUpper().Contains(".PDF"))
                                {
                                    //pdfPath = SPContext.Current.Web.Url + "/" + _objInspecRepository.Title.ToString() + "/" + "SE-" + lblID.Text + "/" + item.Value;
                                    ListPdffiles.Add(item.Value);
                                }
                                #region Bind Images attachments
                                if (item.Value.ToUpper().Contains(".JPG") || item.Value.ToUpper().Contains(".PNG"))
                                {
                                    doc.NewPage();
                                    //string _output = string.Empty;
                                    //lblmsg.Text += "objweb.Title: " + objweb.Title + " item.Value " + item.Value + "<br/>"; continue;//return;
                                    SPFile _itemAttachment = objweb.GetFile(item.Value);

                                    using (Stream s = _itemAttachment.OpenBinaryStream())
                                    {
                                        // Now image in the pdf file
                                        iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(s);
                                        //Resize image depend upon your need
                                        jpg.ScaleToFit(350f, 300f);
                                        //Give space before image
                                        jpg.SpacingBefore = 30f;
                                        //Give some space after the image
                                        jpg.SpacingAfter = 1f;
                                        jpg.Alignment = Element.ALIGN_CENTER;
                                        doc.Add(jpg); //add an image to the created pdf document
                                    }
                                }
                                #endregion
                            }
                        }
                    }
                }
            }
                        
            if (intimagecount > 0)
            {
                doc.Close();
                //ListPdffiles.Add(strpathImages1);
            }
            //if (IRPagescount > 2) ListPdffiles.Add(strpathSplit2);
            #endregion

            #region Apply e-Signature and Comments and Watermark and digtal Signature

            #region Merge all PDFs
            byte[] FinalContent;
            byte[] MergedAtt = MergePDFs(ListPdffiles);

            //byte[] MergedAtt = null;
            using (MemoryStream outputPdfStream = new MemoryStream())//Stream outputPdfStream = new FileStream(filename3, FileMode.Create))
            {
                Document document = new Document();
                PdfSmartCopy copy = new PdfSmartCopy(document, outputPdfStream);
                document.Open();
                PdfReader.unethicalreading = true;
                PdfReader reader;

                int n;
                for (int i = 0; i < ListPdffiles.Count; i++)
                {
                    //byte[] bytes = System.IO.File.ReadAllBytes(pdfArrayFiles[i]);
                    //Response.Write(ListPdffiles[i] + "<br/>");
                    byte[] bytes = objweb.GetFile(ListPdffiles[i]).OpenBinary();
                    //reader = new PdfReader(ListPdffiles[i]);
                    reader = new PdfReader(bytes);
                    

                    n = reader.NumberOfPages;
                    for (int page = 0; page < n; )
                    {
                        copy.AddPage(copy.GetImportedPage(reader, ++page));
                    }
                    reader.Close();
                    reader.Dispose();
                }
                //if (ListPdffiles.Count > 0)
                //{
                    document.Close();
                    document.Dispose();
                    copy.Close();
                    copy.Dispose();
                    MergedAtt = outputPdfStream.GetBuffer();
                    //outputPdfStream.Close();
                //}
            }
            #endregion

            #region CoverPage and comments location
            string CommentsLocation = rblist_CommentsLoc.SelectedValue;
            if (rblist_CommentsLoc.SelectedItem == null)
            {
                lblmsg.Text = "Please choose code and comments placement options.";
                return;
            }

            string FormName = "", EmailBody = "", EmailSubject = "";
            string BldgType = (RIWItem["InspType"] != null) ? RIWItem["InspType"].ToString().Split('#')[1] : "";
            if (BldgType == "Buildings")
                FormName = "RIW_BLDG";
            else FormName = "OTHER_RIW";


            GenericEmailTemplate(objweb, FormName, RIWItem, ref EmailBody, ref EmailSubject);
            string FormBody = EmailBody; //htmlforPdfFile(objweb, LeadItem);
            byte[] CRSFrom_PDFByte = PDFConvert(FormBody);
            FinalContent = MergeBytePDFs(CRSFrom_PDFByte, MergedAtt);

            //if (CommentsLocation == "CRSForm")
            //{
            //    string FormBody = htmlforPdfFile(objweb, LeadItem);
            //    byte[] CRSFrom_PDFByte = PDFConvert(FormBody);
            //    FinalContent = MergeBytePDFs(CRSFrom_PDFByte, MergedAtt);
            //} 
            //else if (CommentsLocation == "PDFAttachment")
            //{
            //    FinalContent = ApplyESignatureWithComments(MergedAtt);
            //}
            //else 
            //{
            //    if (CommentsLocation == "BlankPage")
            //    {
            //        BlankCommentsForm = true;
            //        CommentsPage = GetPDFEmptyPage();
            //        CommentsPage = ApplyESignatureWithComments(CommentsPage);
            //    }
            //    else if (CommentsLocation == "CoverPageFromRIW")
            //    {
            //        CommentsPage = GetRIWCoverPage();
            //        CommentsPage = ApplyESignatureWithComments(CommentsPage);
            //    } 
            //    FinalContent = MergeBytePDFs(CommentsPage, MergedAtt);
            //}
            #endregion

            if (intimagecount > 0) FinalContent = MergeBytePDFs(FinalContent, BindImagesFileStream.GetBuffer());

            if (ckb_ApplyWaterMark.Checked) FinalContent = ApplyWatermark(FinalContent);

            //if (ckb_ApplyDigiSign.Checked)//ApplyDigitalSign 
            //{
            //    string pfxFilePath = ConfigurationManager.AppSettings["pfxFilePath"];
            //    string pfxFilePassword = ConfigurationManager.AppSettings["pfxFilePassword"];
            //    FinalContent = DarPDFSign.SignDocument(pfxFilePath, pfxFilePassword, FinalContent, "RIW Review", "DarAl Handasah", "Jeddah", true, 0, 0, 0, 0);
            //}
            #endregion

            #region Attach Response To RIW Lead task and update
            string attachName = StrFullRef + "-Commented.pdf";
            //Delete old Attachement
            if (LeadItem.Attachments.Count > 0)
            {
                try{LeadItem.Attachments.DeleteNow(attachName);}catch { }
            }

            LeadItem["PartComments"] = PartComments.Text;
            LeadItem["Comment"] = CommentBox.Text;
            LeadItem["Code"] = ddlcode.SelectedValue;

            LeadItem.Attachments.Add(attachName, FinalContent); //File.ReadAllBytes(strpath3)
            objweb.AllowUnsafeUpdates = true;
            LeadItem.Update();// SystemUpdate(false);

            HyperLinkLeadAttach.NavigateUrl = LeadItem.Attachments.UrlPrefix + attachName;
            HyperLinkLeadAttach.Text = attachName;// "Download Compiled RIW Document";

            embedpdfViewer.Visible = true;
            embedpdfViewer.Attributes.Add("src", HyperLinkLeadAttach.NavigateUrl + "#page=1&zoom=150");

            lblmsg.Text = "Success. Please review the compiled response previewed above before submiting and closing the E-Inspection Request.";

            //Response.Write("<script>window.open(\"" + LeadItem.Attachments.UrlPrefix + attachName + "\");</script>");
            //Response.Clear();
            //Response.ContentType = "application/pdf";
            //Response.AppendHeader("Content-Disposition", "inline; filename=" + attachName);
            //Response.TransmitFile(LeadItem.Attachments.UrlPrefix + attachName);
            //Response.End(); 
            #endregion

            SubmitCloseIR.Visible = true;
            #endregion

        }
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail("btnUpdateAREMergePdf_Click Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> Reference:" + lblIRNo.Text + "<br/> WebUrl: " + objweb.Url + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        finally { objweb.Dispose(); }
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

    protected void grdWorkflowtasks_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                CheckBox CheckBoxWTask = (CheckBox)e.Row.FindControl("CheckBoxWTask");
                CheckBoxList chk = (CheckBoxList)e.Row.FindControl("Chkattachments");
                Label lblattachments = (Label)e.Row.FindControl("lblattachments");
                Label lblAssignedTo = (Label)e.Row.FindControl("lblAssignedTo");
                Label lblStatus = (Label)e.Row.FindControl("lblStatus");
                Label lblID = (Label)e.Row.FindControl("lblID");
                Label lblAttachmentsLink = (Label)e.Row.FindControl("lblAttachmentsLink");

                //Label lblSeniorEngineer = (Label)e.Row.FindControl("lblSeniorEngineer");
                Panel PanelFileupload = (Panel)e.Row.FindControl("PanelFileupload");

                Label lblCMQuantity = (Label)e.Row.FindControl("CMQuantity");
                Label lblCMUnit = (Label)e.Row.FindControl("CMUnit");


                SPQuery _WTQuery = new SPQuery();
                _WTQuery.Query = "<Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + lblID.Text + "</Value></Eq></Where>";
                _WTQuery.ViewAttributes = "Scope=\"Recursive\"";
                SPListItemCollection coll = objWorkflowTask.GetItems(_WTQuery);
                SPListItem PartItem = coll[0];

                for (int i = 0; i < PartItem.Attachments.Count; i++)
                {
                    string AttachURL = PartItem.Attachments.UrlPrefix + PartItem.Attachments[i];
                    string filename = PartItem.Attachments[i];
                    string extension = Path.GetExtension(AttachURL);
                    if (extension.ToLower() ==".pdf")
                    {
                        chk.Items.Add(new System.Web.UI.WebControls.ListItem(filename, AttachURL));
                        lblattachments.Text += "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + SPContext.Current.Web.Url + "/" + layoutPath + "/Images/Pdfthumb.png" + " width='50px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                    }
                    else if (extension.ToLower() == ".jpg" || extension.ToLower() == ".png")
                    {
                        chk.Items.Add(new System.Web.UI.WebControls.ListItem(filename, AttachURL));
                        lblattachments.Text += "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + AttachURL.Replace(" ", "%20") + " width='100px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                    }                    
                    else // (!item.Url.ToUpper().Contains(".PDF"))
                    {
                        //chk.Items.Add("" + filename + "");
                        //lblAttachmentsLink.Text = AttachURL;
                        chk.Items.Add(new System.Web.UI.WebControls.ListItem(filename, AttachURL));
                        lblattachments.Text += "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + AttachURL.Replace(" ", "%20") + " width='100px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                    }
                }

                return;
            });
        }
    }

    #region Utilities
    private bool IsUserMemberOfPMCGroup(bool isMember)
    {
        SPSecurity.RunWithElevatedPrivileges(delegate
        {
            using (SPWeb web = SPContext.Current.Site.OpenWeb())
            {
                isMember = web.IsCurrentUserMemberOfGroup(web.Groups[GetParameter("ApproverGroup")].ID);
            }
        });
        return isMember;
    }

    public byte[] GetRIWCoverPage()
    {
        //Get RIW attachment byte
        using (MemoryStream pdfStreamOut = new MemoryStream())
        {
            SPFile spfileAttachment = SPContext.Current.Web.GetFile(RIWItem.Attachments.UrlPrefix + RIWItem.Attachments[0]);
            using (Stream stream = spfileAttachment.OpenBinaryStream())
            {
                PdfReader.unethicalreading = true;
                PdfReader reader = new PdfReader(stream);
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, pdfStreamOut);
                document.Open();
                copy.AddPage(copy.GetImportedPage(reader, 1));
                document.Close();

                return pdfStreamOut.GetBuffer();
            }
        }
    }

    public byte[] GetPDFEmptyPage()
    {
        using (MemoryStream pdfStreamOut = new MemoryStream())
        {
            Document doc = new Document(PageSize.A4, 0, 0, 70, 60);
            PdfWriter writer = PdfWriter.GetInstance(doc, pdfStreamOut);
            doc.Open();
            doc.NewPage();
            //PdfContentByte cb = writer.DirectContent;
            //PdfReader reader = new PdfReader(src);
            //PdfStamper stamper = new PdfStamper(reader, pdfStreamOut);

            // Empty image in the pdf file
            iTextSharp.text.Image jpg = iTextSharp.text.Image.GetInstance(@"C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\14\TEMPLATE\IMAGES\brand-watermark.png");
            //Resize image depend upon your need
            //jpg.ScaleToFit(350f, 300f);
            //Give space before image
            jpg.SpacingBefore = 30f;
            //Give some space after the image
            jpg.SpacingAfter = 1f;
            jpg.Alignment = Element.ALIGN_CENTER;
            doc.Add(jpg); //add an image to the created pdf document

            doc.Close();
            return pdfStreamOut.GetBuffer();
        }
    }
    public byte[] MergeBytePDFs(byte[] PDF1, byte[] PDF2)
    {
        using (MemoryStream outputPdfStream = new MemoryStream())
        {
            Document document = new Document();
            PdfSmartCopy copy = new PdfSmartCopy(document, outputPdfStream);
            document.Open();
            PdfReader.unethicalreading = true;
            PdfReader reader;

            reader = new PdfReader(PDF1);
            for (int page = 0; page < reader.NumberOfPages; )
                copy.AddPage(copy.GetImportedPage(reader, ++page));
            reader.Close();
            reader.Dispose();

            reader = new PdfReader(PDF2);
            for (int page = 0; page < reader.NumberOfPages; )
                copy.AddPage(copy.GetImportedPage(reader, ++page));
            reader.Close();
            reader.Dispose();

            document.Close();
            document.Dispose();
            copy.Close();
            copy.Dispose();
            return outputPdfStream.GetBuffer();
        }
    }
    public byte[] MergePDFs(List<String> pdfArrayFiles)
    {
        ////****************Elevate privelage by impersonating Pool Account
        //WindowsImpersonationContext ctx = null;
        //if (!WindowsIdentity.GetCurrent().IsSystem) ctx = WindowsIdentity.Impersonate(System.IntPtr.Zero);


        ////File.Create(Server.MapPath("~/Reports/payslip4.PDF"), result.Length);
        using (MemoryStream outputPdfStream = new MemoryStream())//Stream outputPdfStream = new FileStream(filename3, FileMode.Create))
        {
           SPSecurity.RunWithElevatedPrivileges(delegate()
           {
                //select the PDF files you want to merge, in this example I only merged 2 files, but you can do more.
                //string[] pdfArrayFiles = { filename1, filename2 };
                Document document = new Document();
                PdfSmartCopy copy = new PdfSmartCopy(document, outputPdfStream);

                document.Open();
                PdfReader.unethicalreading = true;
                PdfReader reader;
                int n;
                for (int i = 0; i < pdfArrayFiles.Count; i++)
                {
                    //byte[] bytes = System.IO.File.ReadAllBytes(pdfArrayFiles[i]);

                    reader = new PdfReader(pdfArrayFiles[i]); n = reader.NumberOfPages;
                    for (int page = 0; page < n; )
                    {
                        copy.AddPage(copy.GetImportedPage(reader, ++page));
                    }
                    reader.Close();
                    reader.Dispose();
                }
               //if (pdfArrayFiles.Count > 0)
               //{
                   document.Close();
                   document.Dispose();
                   copy.Close();
                   copy.Dispose();
              //}
           });
            return outputPdfStream.GetBuffer();
            //outputPdfStream.Close();
        }
        //if (ctx != null) ctx.Undo();// Undo Elevation
        ////**************** End Elevate privelage
    }

    public byte[] ApplyESignatureWithComments(byte[] pdfContent)
    {
        //Deafault Positions that match the Infopath form
        float SigImgX = 250, SigImgY = 120,  UserNameX = 84f, UserNameY = 140f,  DateX = 430f, DateY = 140f,  CodeX = 300f, CodeY= 173f;
        float llx = 65f, lly = 145f, urx = 520f, ury = 235f; //Commments Box lowerLeft and UpperRight points

        //If Blank form is used change Values
        if (BlankCommentsForm)
        {
            SigImgX = 250; SigImgY = 120; 

            UserNameX = 84f; UserNameY = 140f;

            DateX = 430f; DateY = 140f;

            CodeX = 300f; CodeY= 173f;

            llx = 65f; lly = 145f;
            urx = 520f; ury = 235f;
        }

        SPList _objListPFXFile = SPContext.Current.Web.Lists["ESignatures"];
        SPQuery _queryPFX = new SPQuery();
        _queryPFX.Query = "<Where><Eq><FieldRef Name='User' /><Value Type='User'>" + LeadName + "</Value></Eq></Where>";//SPContext.Current.Web.CurrentUser.Name.ToString()
        _queryPFX.ViewAttributes = "Scope=\"Recursive\"";
        SPListItemCollection _PFXItemColl = _objListPFXFile.GetItems(_queryPFX);
        if (_PFXItemColl.Count == 0) throw new Exception("Your Digital Signature is not available in the system.Please contact your PWS administrator."); 
        if (_PFXItemColl.Count > 1)  throw new Exception("Multiple user signature. Please contact your PWS Administrator.");

        Stream inputImageStream = _PFXItemColl[0].File.OpenBinaryStream();

        //SPFile _itemAttachment = SPContext.Current.Web.GetFile(strUrlwithoutattachment);
        //using (Stream streamAttachment = _itemAttachment.OpenBinaryStream())
        using (MemoryStream outputPdfStream = new MemoryStream())//Stream outputPdfStream = new FileStream(strresultPDF, FileMode.Create, FileAccess.Write, FileShare.None))
        {
            var reader = new PdfReader(pdfContent);//streamAttachment);
            var stamper = new PdfStamper(reader, outputPdfStream);
            var pdfContentByte = stamper.GetOverContent(1);

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance((inputImageStream));

            image.SetAbsolutePosition(SigImgX, SigImgY);
            image.ScaleAbsolute(110f, 40f);
            image.Alignment = Element.ALIGN_RIGHT;

            Font BlueFont = FontFactory.GetFont("Arial", 10, Font.NORMAL, BaseColor.BLUE);
            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_MIDDLE, new Phrase(LeadName, BlueFont), UserNameX, UserNameY, 0); //SPContext.Current.Web.CurrentUser.Name
            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_MIDDLE, new Phrase(DateClosed.ToString("dd-MM-yyyy"), BlueFont), DateX, DateY, 0); //System.DateTime.Now.ToString("dd-MM-yyyy")
            pdfContentByte.AddImage(image);

            //string Code = (item["Code"] != null) ? item["Code"].ToString() : ""; //"Approved as noted";//
            //string Comments = (item["Comment"] != null) ? item["Comment"].ToString().Replace("<p>", "").Replace("</p>", "") : "Refer to attached comments";
            string Code = ddlcode.SelectedValue.ToString();// (item["Code"] != null) ? item["Code"].ToString() : ""; //"Approved as noted";//
            string Comments = (CommentBox.Text != null) ? CommentBox.Text.Replace("<p>", "").Replace("</p>", "") : "Refer to attached comments";

            //add text on pdf     
            Font BlueFontText = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLUE);
            ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_MIDDLE, new Phrase(Code, BlueFontText), CodeX, CodeY, 0);//"Status: " + //70f, 220f

            //if (Comments.Length > 80) Comments = Comments.Insert(80, Environment.NewLine + " ");
            ColumnText ct = new ColumnText(pdfContentByte);
            //Rectangle rect = new Rectangle(70f, 210f, 570f, 290f);
            ct.SetSimpleColumn(llx,  lly,  urx,  ury );//(rect.GetLeft(0), 0, rect.GetRight(0), rect.GetTop(0));
            ct.SetText(new Paragraph(Comments, BlueFont)); 
            ct.Leading = 10; //line spacing
            ct.Go();
            
            stamper.Close();
            inputImageStream.Close();
            return outputPdfStream.GetBuffer();
        }
    }

    public byte[] ApplyWatermark(byte[] pdfContent)
    {

        #region get Sig Image from Library

        SPList _objListPFXFile = SPContext.Current.Web.Lists["ESignatures"];
        SPQuery _queryPFX = new SPQuery();
        _queryPFX.Query = "<Where><Eq><FieldRef Name='User' /><Value Type='User'>" + LeadName + "</Value></Eq></Where>";
        _queryPFX.ViewAttributes = "Scope=\"Recursive\"";
        SPListItemCollection _PFXItemColl = _objListPFXFile.GetItems(_queryPFX);
        if (_PFXItemColl.Count == 0) throw new Exception("e-Signature file for " + LeadName + " not available in the system. Please contact PWS administrator.");
        if (_PFXItemColl.Count > 1) throw new Exception("Multiple e-Signature files found" + LeadName + ". Please contact PWS Administrator.");
        Stream inputImageStream = _PFXItemColl[0].File.OpenBinaryStream();
        #endregion

        #region prepare Sig Image
        iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance((inputImageStream));
        image.SetAbsolutePosition(0, 10);
        image.ScaleAbsolute(110f, 40f);
        image.Alignment = Element.ALIGN_RIGHT;
        #endregion
        Font BlueFont = FontFactory.GetFont("Arial", 10, Font.NORMAL, BaseColor.BLUE);
        //Font BlueFontText = FontFactory.GetFont("Arial", 12, Font.NORMAL, BaseColor.BLUE);

        string repeatText = "Reviewed by " + LeadName + " on " +  DateClosed.ToString("dd-MM-yyyy");// System.DateTime.Now.ToString("dd-MM-yyyy");
        using (MemoryStream outputPdfStream = new MemoryStream())
        {
            var reader = new PdfReader(pdfContent);
            var stamper = new PdfStamper(reader, outputPdfStream);

            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                //Get page
                var pdfContentByte = stamper.GetOverContent(i);

                //Add signature
                if (i != 1)
                    pdfContentByte.AddImage(image);

                //add text
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_MIDDLE, new Phrase(repeatText, BlueFont), 110f, 10f, 0);
                ColumnText.ShowTextAligned(pdfContentByte, Element.ALIGN_RIGHT, new Phrase(" Page " + i.ToString() + " of " + reader.NumberOfPages.ToString(), BlueFont), 570f, 10f, 0);
            }

            stamper.Close();
            inputImageStream.Close();
            return outputPdfStream.GetBuffer();
        }
    }

    public string GetParameter(string key)
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
            spquery.ViewAttributes = "Scope=\"Recursive\"";
            SPListItemCollection colitems = list.GetItems(spquery);
            if (colitems.Count == 1)
            {
                SPListItem item = colitems[0];
                if (item["Value"] != null && item["Value"].ToString() != "")
                    retValue = item["Value"].ToString().Trim();
            }
            return retValue;

        }
        catch
        {
            //SendErrorEmail("SiteUrl: <br/>" + web.Url + "<br/><br/>" +
            //               "Error Message:<br/> Error in GetParameter Type:" + key + " <br/><br/> " +
            //               "Message: <br/>" + e.Message + "<br/><br/>" +
            //               "StackTrace: <br/>" + e.StackTrace);
            return "";
        }
        finally
        {
            //web.Dispose();
        }
    }

    public void SendErrorEmail(string Body, string Subject)
    {
        SmtpClient myClient = new SmtpClient(SmtpServer);

        //if (!string.IsNullOrEmpty(GetParameter("FromErrorSmtpUser")))
        //    myClient.Credentials = new System.Net.NetworkCredential(GetParameter("FromErrorSmtpUser"), GetParameter("FromErrorSmtpPWD"));
        //else
        myClient.UseDefaultCredentials = true;
     //   myClient.Credentials = new System.Net.NetworkCredential("darbeirut\\rmzeitouny", "");

        MailMessage msg = new MailMessage();
        msg.From = new MailAddress("PWS_Administrator@dargroup.com", "PWS_Administrator");
        msg.To.Add("Bilal.mustafa@dar.com");
        //msg.To.Add("Bilal.mustafa@dar.com");
        msg.Subject = Environment.MachineName + "-" + Subject;
        msg.Body = Body;
        msg.IsBodyHtml = true;
        try
        {
            myClient.Send(msg);
        }
        catch { }
    }

    public static int GetSPVersion()
    {
        int Val = 0;

        string version = SPFarm.Local.BuildVersion.ToString();

        string[] stateModulesList = version.Split('.');
        version = stateModulesList[0];

        if (int.TryParse(version, out Val))
            Val = int.Parse(version);

        return Val;
    }

    #endregion

    public void MessageAndRedirect(string message, string url)
    {
        Response.Write("<script>alert('" + message + "');</script>");
        if(Referrer != "")
            Response.Write("<script>window.location = '" + Referrer + "';</script>");
        else
            Response.Write("<script>window.location = '" + url + "';</script>");
    }

    public void MessageAndRedirectWithoutAlert(string message, string url)
    {
        //Response.Write("<script>alert('" + message + "');</script>");
        if (Referrer != "")
            Response.Write("<script>window.location = '" + Referrer + "';</script>");
        else
            Response.Write("<script>window.location = '" + url + "';</script>");
    }

    protected void btnSubmitCloseIR_Click(object sender, EventArgs e)
    {
        DateTime Start = DateTime.Now;
        string TimeLog = "";
        try
        {
            if (LeadItem.Attachments.Count ==0)
            {
                lblmsg.Text = "No Compiled Document attachements found on Lead task";
                return;
            }
            else if (LeadItem.Attachments.Count > 1)
            {
                lblmsg.Text = "Multiple attachements on Lead task found, only one final attachment required";
                return;
            }
            else if (LeadItem["Code"] == null)
            {
                lblmsg.Text = "Code is empty ";
                return;
            }
            else
            {
                LeadItem["PartComments"] = PartComments.Text;
                LeadItem["Comment"] = CommentBox.Text;
                LeadItem["Code"] = ddlcode.SelectedValue;
                LeadItem["Status"] = "Completed";
                LeadItem.Update();

                MessageAndRedirectWithoutAlert("Closed successfully", SPContext.Current.Web.Url + redirectURL);
            }
            int total = (int)DateTime.Now.Subtract(Start).TotalSeconds;
            if (total > 5)
            {
                TimeLog += "TOTAL TIME = " + total;
                SendErrorEmail("btnSubmitCloseIR_Click = " + TimeLog, "RIW Lead Submit");
            }

        }
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator";
            SendErrorEmail("btnSubmitCloseIR_Click Error: IR No:" + lblIRNo.Text + "<br/>" + ex.ToString(), ErrorSubject);
        }
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        //SPWeb objweb = null;

        try
        {
            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            //    using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
            //    {
            //        objweb = site.OpenWeb();
            //    }
            //});

            //objweb.AllowUnsafeUpdates = true;
           

            lblmsg.Text = "";
            if (ddlcode.SelectedValue == "Select Code")
            {
                lblmsg.Text = "Not saved, Please select code first.";
                return;
            }
            if (CommentBox.Text == "")
            {
                lblmsg.Text = "Not saved, Please enter final comments.";
                return;
            }

            LeadItem["PartComments"] = PartComments.Text;
            LeadItem["Comment"] = CommentBox.Text;
            LeadItem["Code"] = ddlcode.SelectedValue;
            LeadItem.Update();  //SystemUpdate(false);

            lblmsg.Text = "Saved successfully.";
            btnUpdateAREMergePdf.Visible = true;
        }
        catch (Exception ex) {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail("Page Load Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> RIW Ref:" + lblIRNo.Text + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        //finally { if (objweb != null) objweb.Dispose(); }
    }

    protected void btnCancel_Click(object sender, EventArgs e)
    {
        if (Referrer != "")
            Response.Write("<script>window.location = '" + Referrer + "';</script>");
        else
            Response.Write("<script>window.location = '" + SPContext.Current.Web.Url +"';</script>");
        //SPUtility.Redirect(SPContext.Current.Web.Url, SPRedirectFlags.Default, HttpContext.Current);
    }

    protected void btnAssignPart_Click(object sender, EventArgs e)
    {
        string queryStr = "";

        if (!String.IsNullOrEmpty(RIWQueryStrID)) queryStr = "RIWID=" + RIWQueryStrID;
        else
            if (!String.IsNullOrEmpty(LeadQueryStrID)) queryStr = "ID=" + LeadQueryStrID;

        Response.Write("<script>window.location = '" + SPContext.Current.Web.Url + "/_layouts/PCW/General/EForms/RIWSingleAssign.aspx?" + queryStr + "';</script>");
    }

    #region EMAIL TEMPLATE
    public void GenericEmailTemplate(SPWeb web, string EmailName, SPListItem Listitem, ref string EmailBody, ref string EmailSubject)
    {
        #region Extract Email Body, Subject, ListName From Emails List
        SPList Emails = web.Lists["Emails"];
        SPQuery EmailQuery = new SPQuery();
        EmailQuery.Query = "<Where>" +
                               "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + EmailName + "</Value></Eq>" +
                           "</Where>";
        EmailQuery.ViewAttributes = "Scope=\"Recursive\"";
        SPListItemCollection items = Emails.GetItems(EmailQuery);
        SPListItem item = items[0];
        EmailSubject = (item["EmailSubject"] != null) ? item["EmailSubject"].ToString() : "";
        EmailBody = (item["Body"] != null) ? item["Body"].ToString() : "";
        string ListName = (item["ListName"] != null) ? item["ListName"].ToString() : "";
        Dictionary<string, string> EmailData = new Dictionary<string, string>();
        EmailData.Add("Subject", EmailSubject);
        EmailData.Add("Body", EmailBody);

        string Deliv = GetParameter("UploadDeliverables");
        #endregion


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
                            else
                                                    if (field.Type == SPFieldType.Attachments)
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
                            Logo.ViewAttributes = "Scope=\"Recursive\"";
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
                        else if (ColumnName == "Esign")
                        {
                            #region if Digital Signature Column
                            SPList _AdminPages = web.Lists["ESignatures"];
                            SPQuery Logo = new SPQuery();
                            Logo.Query = "<Where><Eq><FieldRef Name='User' /><Value Type='User'>" + LeadName + "</Value></Eq></Where>";
                            Logo.ViewAttributes = "Scope=\"Recursive\"";
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
                            continue;
                            #endregion
                        }
                        else if (ColumnName == "REEngComment")
                        {
                            value = CommentBox.Text;
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName == "RECode")
                        {
                            bool isInput = false;
                            value = ddlcode.SelectedValue;
                            SetInputTag(inputText, ColumnName, value, ref EmailBody, ref isInput, "");

                            if (!isInput)
                                EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                        }
                        else if (ColumnName == "SentToContractorDate")
                        {
                            #region if E-Inspection Date Closed
                            if (Listitem[ColumnName] != null)
                                value = ((DateTime)Listitem[ColumnName]).ToString("MMM dd, yyyy");
                            else value = DateTime.Today.ToString("MMM dd, yyyy");
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                            #endregion
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
                            #region SET INSPECTION REVIEWERS TABLE
                            SPList list = web.Lists.TryGetList(PartTaskList);
                            if (list != null)
                            {
                                SPQuery listQuery = new SPQuery();
                                listQuery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + Listitem["FullRef"].ToString() + "</Value></Eq></Where>";
                                listQuery.ViewAttributes = "Scope=\"Recursive\"";
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
                                        string Editor = (listItem["HandledBy"] != null) ? listItem["HandledBy"].ToString() : "";
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
                            continue;
                            #endregion
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
                        else if (ColumnName.ToLower() == "ud-lod-path")
                        {
                            #region if UD-LOD-PATH Column For Reassign Form

                            value = "<a href='" + web.Url + "/" + Deliv + "/" + RLODitem["SubmittalRef"] + "'>Repository Folder</a>";

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
        if (ColType == "Boolean")
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
        return "";
    }

    public void SetMultiChoiceLookups(SPWeb web, SPListItem Listitem, ref string EmailBody, bool isDesc_Choice)
    {
        SPList list = web.Lists.TryGetList("InspPurpose");
        if (isDesc_Choice)
            list = web.Lists.TryGetList("InspDesc");

        if (list != null)
        {
            SPQuery SPlistQuery = new SPQuery();
            string ListQuery = "";

            #region BUILD THE QUERY
            ListQuery = "<Where><And>";
            if (isDesc_Choice)
                ListQuery += "<And>";

            ListQuery += "<Eq><FieldRef Name='InspType' /><Value Type='LookupMulti'>" + Listitem["InspType"].ToString().Split('#')[1] + "</Value></Eq>" +
                         "<Eq><FieldRef Name='Discipline' /><Value Type='LookupMulti'>" + Listitem["Discipline"].ToString().Split('#')[1] + "</Value></Eq>" +
                        "</And>";


            string Purpose_Columnname = "InspPurpose";
            if (isDesc_Choice)
            {
                SPFieldLookupValueCollection InspPurplkp = (Listitem[Purpose_Columnname] != null) ? new SPFieldLookupValueCollection(Listitem[Purpose_Columnname].ToString()) : null;
                if (InspPurplkp != null && InspPurplkp.Count > 0)
                {
                    int CountValues = InspPurplkp.Count;
                    if (CountValues == 1)
                        ListQuery += "<Eq><FieldRef Name='" + Purpose_Columnname + "' /><Value Type='LookupMulti'>" + InspPurplkp[0].LookupValue + "</Value></Eq>";
                    else if (CountValues == 2)
                    {
                        ListQuery += "<Or>";
                        for (int i = 0; i < InspPurplkp.Count; i++)
                        {
                            ListQuery += "<Eq><FieldRef Name='" + Purpose_Columnname + "' /><Value Type='LookupMulti'>" + InspPurplkp[i].LookupValue + "</Value></Eq>";
                        }
                        ListQuery += "</Or>";
                    }
                    else
                    {
                        ListQuery += CreateNested_OR_Operator(InspPurplkp.Count - 1);
                        int j = 0;
                        for (int i = 0; i < InspPurplkp.Count; i++)
                        {
                            if (j < 2)
                            {
                                ListQuery += "<Eq><FieldRef Name='" + Purpose_Columnname + "' /><Value Type='LookupMulti'>" + InspPurplkp[i].LookupValue + "</Value></Eq>";
                                if (j == 1)
                                    ListQuery += "</Or>";
                            }
                            else
                            {
                                ListQuery += "<Eq><FieldRef Name='" + Purpose_Columnname + "' /><Value Type='LookupMulti'>" + InspPurplkp[i].LookupValue + "</Value></Eq></Or>";
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

            SPlistQuery.Query = ListQuery;
            SPlistQuery.ViewAttributes = "Scope=\"Recursive\"";

            //Response.Write("Leadname = " + LeadName + "<br/>");

            SPListItemCollection listItems = list.GetItems(SPlistQuery);
            if (listItems.Count > 0)
            {
                SPFieldLookupValueCollection LookupColumn = (Listitem[Purpose_Columnname] != null) ? new SPFieldLookupValueCollection(Listitem[Purpose_Columnname].ToString()) : null;
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
                        for (int i = 0; i < LookupColumn.Count; i++)
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
                    value += "</table>";
                }
            }
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
    #endregion
}


namespace iTextSharpSign
{

    public static class DarPDFSign
    {
        //public static string pfxFilePath = ConfigurationManager.AppSettings["pfxFilePath"];
        //public static string pfxFilePassword = ConfigurationManager.AppSettings["pfxFilePassword"];
        public static byte[] UnSignAndDecryptDocument(byte[] pdfData, string EncryptPass)
        {

            PdfReader reader = new PdfReader(pdfData, new System.Text.ASCIIEncoding().GetBytes(EncryptPass));

            using (MemoryStream stream = new MemoryStream())
            {
                PdfStamper stamper = new PdfStamper(reader, stream);
                //stamper.AcroFields.ClearSignatureField("Signature1");

                stamper.AcroFields.RemoveField("Signature1");

                stamper.Close();
                reader.Close();
                //File.WriteAllBytes(outputFile, stream.ToArray());
                return stream.ToArray(); //GetBuffer(); //
            }
        }

        public static byte[] SignAndEncryptDocument(string pfxFilePath, string pfxFilePassword, byte[] pdfData, string SigReason, string SigContact, string SigLocation, string EncryptPass)
        {
            return SignDocument(pfxFilePath, pfxFilePassword, pdfData, SigReason, SigContact, SigLocation, false, 0, 0, 0, 0, false, false, EncryptPass);

            //Cert myCert = new Cert(pfxFilePath, pfxFilePassword);
            //using (MemoryStream stream = new MemoryStream())
            //{

            //    PdfReader reader = new PdfReader(pdfData);//EncStream.ToArray());//
            //    //To Activate MultiSignatures uncomment this line
            //    //PdfStamper st = PdfStamper.CreateSignature(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0', null, true);
            //    //To disable Multi signatures uncomment this line : every new signature will invalidate older ones !
            //    PdfStamper stamper = PdfStamper.CreateSignature(reader, stream, '\0');//(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0');

            //    //Encrypt with certificate where only users who have the certificate installed can view
            //    //stamper.SetEncryption(myCert.Chain, new int[] { PdfWriter.ALLOW_PRINTING }, PdfWriter.ENCRYPTION_AES_256);
            //    stamper.SetEncryption(PdfWriter.ENCRYPTION_AES_256, "", EncryptPass, PdfWriter.ALLOW_PRINTING);
            //    stamper.SetFullCompression();
            //    //PdfEncryptor.Encrypt(EncReader, EncStream, true, "", EncryptPass, PdfWriter.ALLOW_PRINTING);

            //    PdfSignatureAppearance sAppearance = stamper.SignatureAppearance;
            //    sAppearance.Reason = SigReason;
            //    sAppearance.Contact = SigContact;
            //    sAppearance.Location = SigLocation;

            //    IExternalSignature signature = new PrivateKeySignature(myCert.Akp, "SHA-256");
            //    MakeSignature.SignDetached(sAppearance, signature, myCert.Chain, null, null, null, 0, CryptoStandard.CMS);

            //    stamper.Close();

            //    //return stream.GetBuffer();
            //    return stream.ToArray();
            //}
        }

        public static byte[] SignDocument(string pfxFilePath, string pfxFilePassword, byte[] pdfData, string SigReason, string SigContact, string SigLocation, bool visible, float llx, float lly, float urx, float ury)
        {
            return SignDocument(pfxFilePath, pfxFilePassword, pdfData, SigReason, SigContact, SigLocation, visible, llx, lly, urx, ury, false, false, null);
        }

        public static byte[] SignDocument(string pfxFilePath, string pfxFilePassword, byte[] pdfData, string SigReason, string SigContact, string SigLocation, bool visible, float llx, float lly, float urx, float ury, bool LockChanges, bool SetFullCompression, string EncryptPass)
        {
            Cert myCert = new Cert(pfxFilePath, pfxFilePassword);
            using (MemoryStream stream = new MemoryStream())
            {
                PdfReader reader = new PdfReader(pdfData);
                //To Activate MultiSignatures uncomment this line
                //PdfStamper st = PdfStamper.CreateSignature(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0', null, true);
                //To disable Multi signatures uncomment this line : every new signature will invalidate older ones !
                PdfStamper stamper = PdfStamper.CreateSignature(reader, stream, '\0');//(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0');

                ////Metadata setting
                //st.MoreInfo = this.metadata.getMetaData();
                //st.XmpMetadata = this.metadata.getStreamedMetaData();

                PdfSignatureAppearance sAppearance = stamper.SignatureAppearance;
                sAppearance.Reason = SigReason;
                sAppearance.Contact = SigContact;
                sAppearance.Location = SigLocation;
                if (visible) sAppearance.SetVisibleSignature(new iTextSharp.text.Rectangle(llx, lly, urx, ury), 1, null); //(100, 100, 250, 150); //20, 10, 170, 60

                if (LockChanges) sAppearance.CertificationLevel = PdfSignatureAppearance.CERTIFIED_NO_CHANGES_ALLOWED;
                if (SetFullCompression) stamper.SetFullCompression();
                if (EncryptPass != null) stamper.SetEncryption(PdfWriter.ENCRYPTION_AES_256, "", EncryptPass, PdfWriter.ALLOW_PRINTING);

                IExternalSignature signature = new PrivateKeySignature(myCert.Akp, "SHA-256");
                MakeSignature.SignDetached(sAppearance, signature, myCert.Chain, null, null, null, 0, CryptoStandard.CMS);

                stamper.Close();
                return stream.GetBuffer();
            }

        }

        public static byte[] EncryptDocument(byte[] Data, PdfReader reader, string pass)
        {
            using (MemoryStream output = new MemoryStream())
            {
                PdfEncryptor.Encrypt(reader, output, true, "", pass, PdfWriter.ALLOW_PRINTING);
                return output.GetBuffer();
            }
        }

        //private static byte[] SignDocument(byte[] pdfData, X509Certificate2 cert)
        //{
        //    using (MemoryStream stream = new MemoryStream())
        //    {
        //        PdfReader reader = new PdfReader(pdfData);
        //        PdfStamper stp = PdfStamper.CreateSignature(reader, stream, '\0');
        //        PdfSignatureAppearance sap = stp.SignatureAppearance;

        //        //Protect certain features of the document  
        //        stp.SetEncryption(null,
        //            Guid.NewGuid().ToByteArray(), //random password  
        //            PdfWriter.ALLOW_PRINTING | PdfWriter.ALLOW_COPY | PdfWriter.ALLOW_SCREENREADERS,
        //            PdfWriter.ENCRYPTION_AES_256);

        //        //Get certificate chain 
        //        Org.BouncyCastle.X509.X509CertificateParser cp = new Org.BouncyCastle.X509.X509CertificateParser();
        //        var certChain = new Org.BouncyCastle.X509.X509Certificate[] { cp.ReadCertificate(cert.RawData) };

        ////        sap.SetCrypto(null, certChain, null, PdfSignatureAppearance.WINCER_SIGNED);

        //        //Set signature appearance 
        //        BaseFont helvetica = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1250, BaseFont.EMBEDDED);
        //        iTextSharp.text.Font font = new iTextSharp.text.Font(helvetica, 12, iTextSharp.text.Font.NORMAL);
        //        sap.Layer2Font = font;
        //        sap.SetVisibleSignature(new iTextSharp.text.Rectangle(415, 100, 585, 40), 1, null);

        //        var dic = new PdfSignature(PdfName.ADOBE_PPKMS, PdfName.ADBE_PKCS7_SHA1);
        //        //Set some stuff in the signature dictionary. 
        //        dic.Date = new PdfDate(sap.SignDate);
        //        dic.Name = cert.Subject;    //Certificate name  
        //        if (sap.Reason != null)
        //        {
        //            dic.Reason = sap.Reason;
        //        }
        //        if (sap.Location != null)
        //        {
        //            dic.Location = sap.Location;
        //        }

        //        //Set the crypto dictionary  
        //        sap.CryptoDictionary = dic;

        //        //Set the size of the certificates and signature.  
        //        int csize = 4096; //Size of the signature - 4K 

        //        //Reserve some space for certs and signatures 
        //        var reservedSpace = new Dictionary<PdfName, int>();
        //        reservedSpace[PdfName.CONTENTS] = csize * 2 + 2; //*2 because binary data is stored as hex strings. +2 for end of field 
        //        sap.PreClose(reservedSpace);    //Actually reserve it  

        //        //Build the signature  
        //        System.Security.Cryptography.HashAlgorithm sha = new System.Security.Cryptography.SHA1CryptoServiceProvider();

        //        var sapStream = sap.GetRangeStream();// RangeStream;
        //        int read = 0;
        //        byte[] buff = new byte[8192];
        //        while ((read = sapStream.Read(buff, 0, 8192)) > 0)
        //        {
        //            sha.TransformBlock(buff, 0, read, buff, 0);
        //        }
        //        sha.TransformFinalBlock(buff, 0, 0);

        //        byte[] pk = SignMsg(sha.Hash, cert, false);

        //        //Put the certs and signature into the reserved buffer  
        //        byte[] outc = new byte[csize];
        //        Array.Copy(pk, 0, outc, 0, pk.Length);

        //        //Put the reserved buffer into the reserved space  
        //        PdfDictionary certificateDictionary = new PdfDictionary();
        //        certificateDictionary.Put(PdfName.CONTENTS, new PdfString(outc).SetHexWriting(true));

        //        //Write the signature  
        //        sap.Close(certificateDictionary);

        //        //Close the stamper and save it  
        //        stp.Close();

        //        reader.Close();

        //        //Return the saved pdf  
        //        return stream.GetBuffer();
        //    }

        //}
        //static public byte[] SignMsg(Byte[] msg, X509Certificate2 signerCert, bool detached)
        //{
        //    //  Place message in a ContentInfo object. 
        //    //  This is required to build a SignedCms object. 
        //    Org.BouncyCastle.Asn1.Pkcs.ContentInfo contentInfo = new Org.BouncyCastle.Asn1.Pkcs.ContentInfo(Org.BouncyCastle.Asn1.DerObjectIdentifier, msg);

        //    //  Instantiate SignedCms object with the ContentInfo above. 
        //    //  Has default SubjectIdentifierType IssuerAndSerialNumber. 
        //    System.Security.Cryptography.Pkcs.SignedCms signedCms = new System.Security.Cryptography.Pkcs.SignedCms(contentInfo, detached);

        //    //  Formulate a CmsSigner object for the signer. 
        //    System.Security.Cryptography.Pkcs.CmsSigner cmsSigner = new System.Security.Cryptography.Pkcs.CmsSigner(signerCert);

        //    // Include the following line if the top certificate in the 
        //    // smartcard is not in the trusted list. 
        //    cmsSigner.IncludeOption = X509IncludeOption.EndCertOnly;

        //    //  Sign the CMS/PKCS #7 message. The second argument is 
        //    //  needed to ask for the pin. 
        //    signedCms.ComputeSignature(cmsSigner, false);

        //    //  Encode the CMS/PKCS #7 message. 
        //    return signedCms.Encode();
        //}
    }

    /// <summary>
    /// This class hold the certificate and extract private key needed for e-signature 
    /// </summary>
    class Cert
    {
        #region Attributes

        private string path = "";
        private string password = "";
        private AsymmetricKeyParameter akp;
        private Org.BouncyCastle.X509.X509Certificate[] chain;

        #endregion

        #region Accessors
        public Org.BouncyCastle.X509.X509Certificate[] Chain
        {
            get { return chain; }
        }
        public AsymmetricKeyParameter Akp
        {
            get { return akp; }
        }

        public string Path
        {
            get { return path; }
        }

        public string Password
        {
            get { return password; }
            set { password = value; }
        }
        #endregion

        #region Helpers

        private void processCert()
        {
            string alias = null;

            Pkcs12Store pk12;

            //First we'll read the certificate file
            pk12 = new Pkcs12Store(new FileStream(this.Path, FileMode.Open, FileAccess.Read), this.password.ToCharArray());

            //then Iterate throught certificate entries to find the private key entry
            IEnumerator i = pk12.Aliases.GetEnumerator();//aliases();
            while (i.MoveNext())
            {
                alias = ((string)i.Current);
                if (pk12.IsKeyEntry(alias))
                    break;
            }

            this.akp = pk12.GetKey(alias).Key; //getKey();
            X509CertificateEntry[] ce = pk12.GetCertificateChain(alias);
            this.chain = new Org.BouncyCastle.X509.X509Certificate[ce.Length];
            for (int k = 0; k < ce.Length; ++k)
                chain[k] = ce[k].Certificate;// getCertificate();

        }
        #endregion

        #region Constructors
        public Cert()
        { }
        public Cert(string cpath)
        {
            this.path = cpath;
            this.processCert();
        }
        public Cert(string cpath, string cpassword)
        {
            this.path = cpath;
            this.Password = cpassword;
            this.processCert();
        }
        #endregion

    }

    /// <summary>
    /// This is a holder class for PDF metadata
    /// </summary>
    class MetaData
    {
        private Hashtable info = new Hashtable();

        public Hashtable Info
        {
            get { return info; }
            set { info = value; }
        }

        public string Author
        {
            get { return (string)info["Author"]; }
            set { info.Add("Author", value); }
        }
        public string Title
        {
            get { return (string)info["Title"]; }
            set { info.Add("Title", value); }
        }
        public string Subject
        {
            get { return (string)info["Subject"]; }
            set { info.Add("Subject", value); }
        }
        public string Keywords
        {
            get { return (string)info["Keywords"]; }
            set { info.Add("Keywords", value); }
        }
        public string Producer
        {
            get { return (string)info["Producer"]; }
            set { info.Add("Producer", value); }
        }

        public string Creator
        {
            get { return (string)info["Creator"]; }
            set { info.Add("Creator", value); }
        }

        public Hashtable getMetaData()
        {
            return this.info;
        }

        //public byte[] getStreamedMetaData()
        //{
        //    MemoryStream os = new System.IO.MemoryStream();
        //    //iTextSharp.text.xml.xmp.XmpWriter
        //    XmpWriter xmp = new XmpWriter(os, (PdfDictionary)this.info);            
        //    xmp.Close();            
        //    return os.ToArray();
        //}

    }


    /// <summary>
    /// this is the most important class
    /// it uses iTextSharp library to sign a PDF document
    /// </summary>
    class PDFSigner
    {
        private string inputPDF = "";
        private string outputPDF = "";
        private Cert myCert;
        private MetaData metadata;

        public PDFSigner(string input, string output)
        {
            this.inputPDF = input;
            this.outputPDF = output;
        }

        public PDFSigner(string input, string output, Cert cert)
        {
            this.inputPDF = input;
            this.outputPDF = output;
            this.myCert = cert;
        }
        public PDFSigner(string input, string output, MetaData md)
        {
            this.inputPDF = input;
            this.outputPDF = output;
            this.metadata = md;
        }
        public PDFSigner(string input, string output, Cert cert, MetaData md)
        {
            this.inputPDF = input;
            this.outputPDF = output;
            this.myCert = cert;
            this.metadata = md;
        }

        public void Verify()
        {

        }

        public void Sign(string SigReason, string SigContact, string SigLocation, bool visible)
        {
            Sign(SigReason, SigContact, SigLocation, visible, 20, 10, 170, 60);
        }
        public void Sign(string SigReason, string SigContact, string SigLocation, bool visible, float llx, float lly, float urx, float ury)
        {
            PdfReader reader = new PdfReader(this.inputPDF);
            //Activate MultiSignatures
            //PdfStamper st = PdfStamper.CreateSignature(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0', null, true);
            //To disable Multi signatures uncomment this line : every new signature will invalidate older ones !
            PdfStamper st = PdfStamper.CreateSignature(reader, new FileStream(this.outputPDF, FileMode.Create, FileAccess.Write), '\0');

            ////Metadata setting
            //st.MoreInfo = this.metadata.getMetaData();
            //st.XmpMetadata = this.metadata.getStreamedMetaData();

            PdfSignatureAppearance sAppearance = st.SignatureAppearance;

            //sap.SetCrypto(this.myCert.Akp, this.myCert.Chain, null, PdfSignatureAppearance.CERTIFIED_NO_CHANGES_ALLOWED);//WINCER_SIGNED);
            sAppearance.Reason = SigReason;
            sAppearance.Contact = SigContact;
            sAppearance.Location = SigLocation;
            //sAppearance.
            if (visible)
                sAppearance.SetVisibleSignature(new iTextSharp.text.Rectangle(llx, lly, urx, ury), 1, null);
            //sap.SetVisibleSignature(new iTextSharp.text.Rectangle(100, 100, 250, 150), 1, null); //20, 10, 170, 60

            IExternalSignature signature = new PrivateKeySignature(this.myCert.Akp, "SHA-256");
            MakeSignature.SignDetached(sAppearance, signature, this.myCert.Chain, null, null, null, 0, CryptoStandard.CMS);
            st.Close();
        }
    }
}


