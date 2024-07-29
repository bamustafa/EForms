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
using System.Text;

public partial class CompileForm : System.Web.UI.Page
{
    #region GLOBAL VARIABLES
    string _ColumnName = "";
    public static SPWeb objweb = null;
    SPList objlist = null, objlistLeadTasks = null, objWorkflowTask = null;
    SPListItem RIWItem = null, LeadItem = null;
    SPFieldUserValue LeadUser = null;

    public string Category = "";
    public string SLF_ITEMS_BODY = "";

    public static string ErrorSubject = "EForms - Error on CompileForm.aspx.cs",
                         SmtpServer = ConfigurationManager.AppSettings["SmtpServer"],
                         layoutPath = "_layouts/15/PCW/General/EForms",
                         LeadRedirectURL = "/Lists/LeadAction/",
                         isDigitalForm = "";

    public static int Version = GetSPVersion();
    
    private string RIWQueryStrID
    {
        get
        {
            return (string)ViewState["RIWQueryStrID"];
        }

        set
        {
            ViewState["RIWQueryStrID"] = value;
        }
    }

    private string LeadQueryStrID
    {
        get
        {
            return (string)ViewState["LeadQueryStrID"];
        }

        set
        {
            ViewState["LeadQueryStrID"] = value;
        }
    }

    private string MainList
    {
        get
        {
            return (string)ViewState["MainList"];
        }

        set
        {
            ViewState["MainList"] = value;
        }
    }

    private string LeadTaskList
    {
        get
        {
            return (string)ViewState["LeadTaskList"];
        }

        set
        {
            ViewState["LeadTaskList"] = value;
        }
    }

    private string PartTaskList
    {
        get
        {
            return (string)ViewState["PartTaskList"];
        }

        set
        {
            ViewState["PartTaskList"] = value;
        }
    }

    private string DeliverableType
    {
        get
        {
            return (string)ViewState["DeliverableType"];
        }

        set
        {
            ViewState["DeliverableType"] = value;
        }
    }

    private string StrFullRef
    {
        get
        {
            return (string)ViewState["StrFullRef"];
        }

        set
        {
            ViewState["StrFullRef"] = value;
        }
    }

    private string Referrer
    {
        get
        {
            return (string)ViewState["Referrer"];
        }

        set
        {
            ViewState["Referrer"] = value;
        }
    }

    private string InspectionReviewCodes
    {
        get
        {
            return (string)ViewState["InspectionReviewCodes"];
        }

        set
        {
            ViewState["InspectionReviewCodes"] = value;
        }
    }

    private string InspectionOutCO
    {
        get
        {
            return (string)ViewState["InspectionOutCO"];
        }

        set
        {
            ViewState["InspectionOutCO"] = value;
        }
    }

    private string InspectionOutDC
    {
        get
        {
            return (string)ViewState["InspectionOutDC"];
        }

        set
        {
            ViewState["InspectionOutDC"] = value;
        }
    }

    private string _IssuedItem_Email
    {
        get
        {
            return (string)ViewState["_IssuedItem_Email"];
        }

        set
        {
            ViewState["_IssuedItem_Email"] = value;
        }
    }

    private string _LeadTaskClose_Email
    {
        get
        {
            return (string)ViewState["_LeadTaskClose_Email"];
        }

        set
        {
            ViewState["_LeadTaskClose_Email"] = value;
        }
    }

    private string FormCodes
    {
        get
        {
            return (string)ViewState["FormCodes"];
        }

        set
        {
            ViewState["FormCodes"] = value;
        }
    }

    private string LeadName
    {
        get
        {
            return (string)ViewState["LeadName"];
        }

        set
        {
            ViewState["LeadName"] = value;
        }
    }

    private string Discipline
    {
        get
        {
            return (string)ViewState["Discipline"];
        }

        set
        {
            ViewState["Discipline"] = value;
        }
    }

    private string UserName
    {
        get
        {
            return (string)ViewState["UserName"];
        }

        set
        {
            ViewState["UserName"] = value;
        }
    }

    private string FormName
    {
        get
        {
            return (string)ViewState["FormName"];
        }

        set
        {
            ViewState["FormName"] = value;
        }
    }

    private string redirectURL
    {
        get
        {
            return (string)ViewState["redirectURL"];
        }

        set
        {
            ViewState["redirectURL"] = value;
        }
    }

    private string _RefColumn
    {
        get
        {
            return (string)ViewState["_RefColumn"];
        }

        set
        {
            ViewState["_RefColumn"] = value;
        }
    }

    private bool ApplyDigitalSign
    {
        get
        {
            return (bool)ViewState["ApplyDigitalSign"];
        }

        set
        {
            ViewState["ApplyDigitalSign"] = value;
        }
    }

    private bool BlankCommentsForm
    {
        get
        {
            return (bool)ViewState["BlankCommentsForm"];
        }

        set
        {
            ViewState["BlankCommentsForm"] = value;
        }
    }

    private bool isCompile
    {
        get
        {
            return (bool)ViewState["isCompile"];
        }

        set
        {
            ViewState["isCompile"] = value;
        }
    }

    private bool isTaskOpen
    {
        get
        {
            if (ViewState["isTaskOpen"] != null)
                return (bool)ViewState["isTaskOpen"];
            else
                return false;
        }

        set
        {
            ViewState["isTaskOpen"] = value;
        }
    }

    private bool IgnorePDFFile
    {
        get
        {            
            if (ViewState["IgnorePDFFile"] != null)
                return (bool)ViewState["IgnorePDFFile"];
            else
                return false;
        }

        set
        {
            ViewState["IgnorePDFFile"] = value;
        }
    }
    
    private bool isMain
    {
        get
        {
            if (ViewState["isMain"] != null)
                return (bool)ViewState["isMain"];
            else
                return false;
        }

        set
        {
            ViewState["isMain"] = value;
        }
    }


    private DateTime DateClosed
    {
        get
        {
            return (DateTime)ViewState["DateClosed"];
        }

        set
        {
            ViewState["DateClosed"] = value;
        }
    }
    #endregion

    #region GENERIC EMAIL TEMPLATE VARIABLES

    string Title = "",
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
            
            string TimeLog = "";
            ApplyDigitalSign = true;
            BlankCommentsForm = false;
            isCompile = true;
            

            if (Request.UrlReferrer != null)
                Referrer = Request.UrlReferrer.ToString();

            LeadName = SPContext.Current.Web.CurrentUser.Name;
            string UserEmail = "";
            try { UserEmail = SPContext.Current.Web.CurrentUser.Email; }
            catch { }

            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    objweb = site.OpenWeb();
                }

                objweb.AllowUnsafeUpdates = true;
            });

            GetSPListItem();

            if (!IsPostBack)
            {
                #region IS ALLOWED TO CLOSE THE TASK 

                if (RIWItem == null)
                {
                    if (LeadItem != null)
                    {
                        MessageAndRedirect("This " + LeadItem["Reference"].ToString() + " has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);
                        SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>Page Load Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> RIW Ref: " + LeadItem["Reference"].ToString(), ErrorSubject);
                    }
                    else
                        MessageAndRedirect("This item has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);

                    return;
                }
                else
                    DateClosed = (RIWItem["SentToContractorDate"] != null) ? (DateTime)RIWItem["SentToContractorDate"] : DateTime.Now;

                if (LeadItem != null)
                {
                    SPFieldUserValueCollection AssignedToLead = (LeadItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(objweb, LeadItem["AssignedTo"].ToString()) : null;
                    bool goAhead = false;
                    if (!SPContext.Current.Web.CurrentUser.IsSiteAdmin)
                    {
                        for (int i = 0; i < AssignedToLead.Count; i++)
                        {
                            if (AssignedToLead[i].User.Email.ToString().ToLower() == UserEmail.ToLower())
                                goAhead = true;
                        }

                        if (!goAhead)
                            MessageAndRedirect("This task is not assigned to you.", SPContext.Current.Web.Url + redirectURL);
                    }
                    #endregion
                    DateTime Start6 = DateTime.Now;

                    DateTime Start6_1 = DateTime.Now;
                    #region Creating Data Table with required columns for Final Output.

                    DataTable dtstore = new DataTable();
                    dtstore.Columns.Add("ID", typeof(string));
                    dtstore.Columns.Add("Title", typeof(string));
                    dtstore.Columns.Add("Trade", typeof(string));
                    dtstore.Columns.Add("AssignedTo", typeof(string));
                    //dtstore.Columns.Add("SeniorEngineer", typeof(string));
                    //dtstore.Columns.Add("PartRE", typeof(string));
                    dtstore.Columns.Add("Date", typeof(DateTime));
                    dtstore.Columns.Add("Status", typeof(string));
                    dtstore.Columns.Add("Code", typeof(string));
                    //dtstore.Columns.Add("CMQuantity", typeof(string));
                    //dtstore.Columns.Add("CMUnit", typeof(string)); 
                    dtstore.Columns.Add("Comment", typeof(string));
                    dtstore.Columns.Add("RejectionTrades", typeof(string));
                    //dtstore.Columns.Add("DaysHeld", typeof(string));
                    //dtstore.Columns.Add("WorkflowLink", typeof(string));

                    DataTable dtFinal = new DataTable();
                    dtFinal.Columns.Add("ID", typeof(string));
                    dtFinal.Columns.Add("Title", typeof(string));
                    dtFinal.Columns.Add("Trade", typeof(string));
                    dtFinal.Columns.Add("AssignedTo", typeof(string));
                    //dtFinal.Columns.Add("SeniorEngineer", typeof(string));
                    //dtFinal.Columns.Add("PartRE", typeof(string));
                    dtFinal.Columns.Add("Date", typeof(DateTime));
                    dtFinal.Columns.Add("Status", typeof(string));
                    dtFinal.Columns.Add("Code", typeof(string));
                    //dtFinal.Columns.Add("CMQuantity", typeof(string));
                    //dtFinal.Columns.Add("CMUnit", typeof(string));
                    dtFinal.Columns.Add("Comment", typeof(string));
                    dtFinal.Columns.Add("RejectionTrades", typeof(string));


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

                    lblIRNo.Text = (RIWItem[_RefColumn] != null) ? RIWItem[_RefColumn].ToString() : "";
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
                    PartComments.Text = (LeadItem["PartComments"] != null) ? LeadItem["PartComments"].ToString().Replace("\n", "<br>") : "";
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
                    if (DeliverableType != "SLF" && DeliverableType != "SCR")
                    {
                        if (RIWItem.Attachments.Count == 0)
                        {
                            //MessageAndRedirect("This RIW has no attachment to get the cover page from, please close it from Lead Task.", SPContext.Current.Web.Url + LeadRedirectURL + "DispForm.aspx?ID=" + LeadItemID.ToString());
                            Response.Write("<script>alert('This E-" + DeliverableType + " has no attachment to get the cover page from, please close it from Lead Task.');</script>");
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
                            lblmsg.Text = "This Inspection has multiple attachments, you must close this Inspection from lead task.";
                            return;
                        }
                    }
                    #endregion
                    TimeLog += "checking the Status - Populate the Inspection info in grid - Page_Load - " + DateTime.Now.Subtract(Start6_2).TotalSeconds + "<br/><br/>";

                    DateTime Start6_3 = DateTime.Now;
                    #region getting the records from Part Tasks

                    SPQuery _WTQuery = new SPQuery();
                    _WTQuery.Query = "<Where><And>" +
                                        "<Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq>" +
                                         "<Or><Or>" +
                                           "<Eq><FieldRef Name='isPMCAssignment' /><Value Type='Boolean'>0</Value></Eq>" +
                                           "<Eq><FieldRef Name='IsLeadAction' /><Value Type='Boolean'>1</Value></Eq>" +
                                         "</Or>" +
                                         "<IsNull><FieldRef Name='Role' /></IsNull>" +
                                         "</Or>" +
                                        "</And></Where>";
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
                                //string _Role = dtTask.Rows[i]["Role"] != null ? dtTask.Rows[i]["Role"].ToString() : "";
                                string _Trade = dtTask.Rows[i]["Trade"].ToString() != "" ? dtTask.Rows[i]["Trade"].ToString() : "SLF Coordinator";
                                string _Status = dtTask.Rows[i]["Status"].ToString();
                                DataRow dr = dtstore.NewRow();
                                dr["ID"] = dtTask.Rows[i]["ID"].ToString();
                                dr["Title"] = dtTask.Rows[i]["Title"].ToString();
                                dr["Trade"] = _Trade;
                                dr["AssignedTo"] = dtTask.Rows[i]["AssignedTo"].ToString();
                                //Get User collection
                                if (dtTask.Rows[i]["AssignedTo"].ToString().Contains(";#"))
                                {
                                    SPListItem taskItem = objWorkflowTask.GetItemById(int.Parse(dtTask.Rows[i]["ID"].ToString()));
                                    SPFieldUserValueCollection usrColl = (taskItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(objweb, taskItem["AssignedTo"].ToString()) : null;
                                    string PartNames = "";
                                    foreach (SPFieldUserValue usrval in usrColl)
                                        if (!PartNames.Contains(usrval.User.Name)) PartNames += usrval.User.Name + "<br/>";
                                    dr["AssignedTo"] = PartNames;
                                }
                                dr["Date"] = dtTask.Rows[i]["Date"];
                                dr["Status"] = _Status;
                                dr["Code"] = dtTask.Rows[i]["Code"].ToString();
                                dr["Comment"] = dtTask.Rows[i]["Comment"].ToString();

                                dr["RejectionTrades"] = "";

                                if (_Status == "Open")
                                    isTaskOpen = true;
                                dtstore.Rows.Add(dr);

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
                    #endregion

                    TimeLog += "Total checking the status of the IR from RIW list. - Page_Load - " + DateTime.Now.Subtract(Start6).TotalSeconds + "<br/><br/>";

                    int total = (int)DateTime.Now.Subtract(Start).TotalSeconds;
                    if (total > 5)
                    {
                        TimeLog += "TOTAL TIME = " + total;
                        //SendErrorEmail(StrFullRef + " - Page_Load<br/><br/>" + TimeLog, "RIW Lead Page Load");
                    }

                    SetButtons();
                }
            }
        }
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>Page Load Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> RIW Ref:" + lblIRNo.Text + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        //finally { if (objweb != null) objweb.Dispose(); }
    }    

    public void GetSPListItem()
    {
        #region QUERY E-INSPECTION AND LEAD
        isDigitalForm = GetParameter(objweb, "isDigitalForm");
        if (string.IsNullOrEmpty(isDigitalForm))
            isDigitalForm = "Yes";

        Guid ListId = new Guid(HttpContext.Current.Request.QueryString["ListId"]);
        loader.ImageUrl = objweb.Url + ConfigurationManager.AppSettings["loaderUrl"];

        RIWQueryStrID = HttpContext.Current.Request.QueryString["ModuleID"];
        LeadQueryStrID = HttpContext.Current.Request.QueryString["ID"];

        int LeadItemID = 0;
        int RIWItemID = 0;
        if (!String.IsNullOrEmpty(RIWQueryStrID))
        {           
            isMain = true;

            objlist = objweb.Lists[ListId];
            DateTime Start1 = DateTime.Now;
            if (int.TryParse(RIWQueryStrID, out RIWItemID))
            {
                //RIWItemID = int.Parse(RIWQueryStrID);

                SPQuery _query = new SPQuery();
                _query.Query = "<Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + RIWItemID + "</Value></Eq></Where>";

                _query.ViewAttributes = "Scope=\"Recursive\"";
                SPListItemCollection RIWItems = objlist.GetItems(_query);
                if (RIWItems.Count == 1)
                    RIWItem = RIWItems[0];
                else
                {
                    MessageAndRedirect("This " + objlist.Title + " has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);
                    return;
                }
                //if (!IsPostBack)
                StrFullRef = (RIWItem["FullRef"] != null) ? RIWItem["FullRef"].ToString() : "";
                Discipline = (RIWItem["Discipline"] != null) ? RIWItem["Discipline"].ToString().Split('#')[1] : "";
                DeliverableType = (RIWItem["DeliverableType"] != null) ? RIWItem["DeliverableType"].ToString() : "";
              
                //redirectURL = LeadRedirectURL;
                redirectURL = Referrer;//"/Lists/" + RIWItem.ParentList.Title;
            }
            //TimeLog += "GET ITEM FROM MAIN " + DeliverableType + " - Page_Load - " + DateTime.Now.Subtract(Start1).TotalSeconds + "<br/><br/>";
        }
        else if (!String.IsNullOrEmpty(LeadQueryStrID))
        {
            //isMain = false;

            //Response.Write("In Lead");

            objlistLeadTasks = objweb.Lists[ListId];
            DateTime Start2 = DateTime.Now;
            if (int.TryParse(LeadQueryStrID, out LeadItemID))
            {
                {
                    //LeadItemID = int.Parse(LeadQueryStrID);
                    LeadItem = objlistLeadTasks.GetItemById(LeadItemID);               

                    StrFullRef = (LeadItem["Reference"] != null) ? LeadItem["Reference"].ToString() : "";
                    Discipline = (LeadItem["Trade"] != null) ? LeadItem["Trade"].ToString() : "";
                    DeliverableType = (LeadItem["DeliverableType"] != null) ? LeadItem["DeliverableType"].ToString() : "";
                    redirectURL = Referrer;//"/Lists/" + LeadItem.ParentList.Title;
                    LeadUser = (LeadItem["AssignedTo"] != null) ? new SPFieldUserValue(objweb, LeadItem["AssignedTo"].ToString()) : null;
                    //string FromStation = (RIWItem["FromStation"] != null) ? RIWItem["FromStation"].ToString().Split('#')[1] : "";
                    //Response.Write("FromStation = " + FromStation);
                    //LeadName = LeadUser.User.Name;// SPContext.Current.Web.CurrentUser.Name;//Used in getting the signature image file

                }
            }
            else //contains text means fullRef
                StrFullRef = HttpContext.Current.Request.QueryString["ID"];
            //TimeLog += "GET ITEM FROM LEAD " + DeliverableType + " - Page_Load - " + DateTime.Now.Subtract(Start2).TotalSeconds + "<br/><br/>";
                        
        }        

        _RefColumn = "FullRef";
        if (DeliverableType == "SLF")
            _RefColumn = "Reference";

        EXTRACT_MAJORTYPE_LIST_VALUES(objweb, DeliverableType);

        if (!IsPostBack)
        {
            if (!string.IsNullOrEmpty(FormCodes))
            {
                string[] codeArray = FormCodes.Split('#');
                ddlcode.Items.Clear();
                ddlcode.Items.Add("Select Code");
                for (int i = 0; i < codeArray.Length; i++)
                {
                    string _code = codeArray[i].Replace(";", "");
                    if (!string.IsNullOrEmpty(_code))
                        ddlcode.Items.Add(_code);
                }
            }
        }

        objlist = objweb.Lists[MainList];
        objlistLeadTasks = objweb.Lists[LeadTaskList];
        objWorkflowTask = objweb.Lists[PartTaskList];

        //Get RIW and Lead Item by Ref
        if (RIWItem == null)
        {
            DateTime Start3 = DateTime.Now;
            SPQuery _query = new SPQuery();
            _query.Query = "<Where><Eq><FieldRef Name='" + _RefColumn + "' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
            _query.ViewAttributes = "Scope=\"Recursive\"";
            //_query.ViewFields = RIWViewFields;
            SPListItemCollection RIWItems = objlist.GetItems(_query);
            if (RIWItems.Count == 1)
                RIWItem = RIWItems[0];
            else
            {
                MessageAndRedirect("This " + DeliverableType + " has issues. Please Contact PWS Administrator.", SPContext.Current.Web.Url + redirectURL);
                return;
            }
            //TimeLog += "GET ITEM FROM MAIN " + DeliverableType + " IF NULL - Page_Load - " + DateTime.Now.Subtract(Start3).TotalSeconds + "<br/><br/>";
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
                MessageAndRedirect("This " + DeliverableType + " has no lead task. Please contact PWS administrator.", SPContext.Current.Web.Url + redirectURL);
                return;
            }
            //TimeLog += "GET ITEM FROM Lead " + DeliverableType + " IF NULL - Page_Load - " + DateTime.Now.Subtract(Start4).TotalSeconds + "<br/><br/>";
        }
        #endregion
    }

    public byte[] PDFConvert(string html)
    {
        PdfConverter pdfConverter = new PdfConverter();
        pdfConverter.AuthenticationOptions.Username = CredentialCache.DefaultNetworkCredentials.UserName;
        pdfConverter.AuthenticationOptions.Password = CredentialCache.DefaultNetworkCredentials.Password;
        pdfConverter.PdfDocumentOptions.PdfPageSize = PdfPageSize.A4;
        //if(DeliverableType == "SLF")
         //pdfConverter.PdfDocumentOptions.PdfPageOrientation = PDFPageOrientation.Landscape;
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
                CheckBox IncludePartCommentCheckBox = (CheckBox)e.Row.FindControl("IncludePartCommentCheckBox");
                CheckBoxList chk = (CheckBoxList)e.Row.FindControl("Chkattachments");
                TextBox txtRejTrades = (TextBox)e.Row.FindControl("txtRejTrades");

                Label lblID = (Label)e.Row.FindControl("lblID");
                Label lblStatus = (Label)e.Row.FindControl("lblStatus");
                Label lblattachments = (Label)e.Row.FindControl("lblattachments");

                SPQuery _WTQuery = new SPQuery();
                _WTQuery.Query = "<Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + lblID.Text + "</Value></Eq></Where>";
                _WTQuery.ViewFields = "<FieldRef Name='DeliverableType' />" +
                                      "<FieldRef Name='Role' />" +
                                      "<FieldRef Name='IsLeadAction' />" +
                                      "<FieldRef Name='Attachments' />";
                SPListItemCollection coll = objWorkflowTask.GetItems(_WTQuery);
                if (coll.Count == 0)
                    return;

                SPListItem PartItem = coll[0];

                string DeliverableType = (PartItem["DeliverableType"] != null) ? PartItem["DeliverableType"].ToString() : "";
                bool IsLeadAction = (PartItem["IsLeadAction"] != null) ? bool.Parse(PartItem["IsLeadAction"].ToString()) : false;

                string Role = (PartItem["Role"] != null) ? PartItem["Role"].ToString() : "";

                if (Role == "TeamLeader" || IsLeadAction)
                {
                    CheckBoxWTask.Enabled = true;
                    txtRejTrades.Enabled = true;
                }
                else
                {
                    CheckBoxWTask.Enabled = false;
                    txtRejTrades.Enabled = false;
                }

                if (DeliverableType == "SCR" || DeliverableType == "MAT" || DeliverableType == "MAT" || DeliverableType == "SLF")
                {
                    string headerTextToHide = "Include Part Comment ?";
                    string AttachmentHeaderText = "Attachments (select to include)";

                    for (int i = 0; i < e.Row.Cells.Count; i++)
                    {
                        DataControlField field = grdWorkflowtasks.Columns[i];

                        if (field.HeaderText == headerTextToHide)
                            field.Visible = false;

                        if (DeliverableType == "SLF" && field.HeaderText == AttachmentHeaderText)
                            field.Visible = false;
                    }
                }
                if (DeliverableType != "SLF")
                {
                    for (int i = 0; i < PartItem.Attachments.Count; i++)
                    {
                        string ImgFileUrl = "", FullFileUrl = "", linkColor = "#6083e0 !important";
                        string AttachURL = PartItem.Attachments.UrlPrefix + PartItem.Attachments[i];
                        string filename = PartItem.Attachments[i];
                        string extension = Path.GetExtension(AttachURL);
                        FullFileUrl = "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0; color:" + linkColor + "' target='_blank' >" + filename + "</a>";
                        if (extension.ToLower() == ".pdf")
                        {
                            ImgFileUrl = "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + SPContext.Current.Web.Url + "/" + layoutPath + "/Images/Pdfthumb.png" + " width='50px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                            chk.Items.Add(new System.Web.UI.WebControls.ListItem(FullFileUrl, filename));
                            lblattachments.Text += ImgFileUrl;
                        }
                        else if (extension.ToLower() == ".jpg" || extension.ToLower() == ".jpeg" || extension.ToLower() == ".png")
                        {
                            ImgFileUrl = "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + AttachURL.Replace(" ", "%20") + " width='100px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                            chk.Items.Add(new System.Web.UI.WebControls.ListItem(FullFileUrl, filename));
                            lblattachments.Text += ImgFileUrl;
                        }
                        else
                        {
                            chk.Items.Add(new System.Web.UI.WebControls.ListItem(FullFileUrl, filename));
                            lblattachments.Text += "<a href=" + AttachURL.Replace(" ", "%20") + " style='border:0;' target='_blank' ><img src=" + AttachURL.Replace(" ", "%20") + " width='100px' height='50px'  alt='' style='border:0;'  /></a><br/>";
                        }
                    }
                }

                if (lblStatus.Text == "Open")
                {
                    CheckBoxWTask.Enabled = false;
                    IncludePartCommentCheckBox.Checked = false;
                    IncludePartCommentCheckBox.Enabled = false;
                    chk.Enabled = false;
                }

                return;
            });
        }
    }

    #region UTILITIES
    public void SetButtons()
    {
        IgnorePDFFile = false;
        if (!string.IsNullOrEmpty(isDigitalForm))
        {
            if (isDigitalForm.ToLower() == "no")
            {
                btnUpdateAREMergePdf.Visible = false;
                SubmitCloseIR.Visible = true;
                FUPanel.Visible = true;
            }
            else
            {
                int countOpenItems = GET_SLF_OPENITEMS_COUNT(objweb, StrFullRef);
                if (countOpenItems > 0)
                {
                    IgnorePDFFile = true;
                    btnUpdateAREMergePdf.Visible = false;
                    SubmitCloseIR.Visible = true;
                }
                else
                {
                    if (IgnorePDFFile) IgnorePDFFile = false;
                    btnUpdateAREMergePdf.Visible = true;
                    if(string.IsNullOrEmpty(HyperLinkLeadAttach.NavigateUrl))
                     SubmitCloseIR.Visible = false;
                    else SubmitCloseIR.Visible = true;
                }
                    FUPanel.Visible = false;
            }
        }
    }
    private bool IsUserMemberOfPMCGroup(bool isMember)
    {
        SPSecurity.RunWithElevatedPrivileges(delegate
        {
            using (SPWeb web = SPContext.Current.Site.OpenWeb())
            {
                isMember = web.IsCurrentUserMemberOfGroup(web.Groups[GetParameter(web, "ApproverGroup")].ID);
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

            if (PDF2 != null)
            {
                reader = new PdfReader(PDF2);
                for (int page = 0; page < reader.NumberOfPages;)
                    copy.AddPage(copy.GetImportedPage(reader, ++page));
                reader.Close();
                reader.Dispose();
            }

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

                    reader = new PdfReader(pdfArrayFiles[i]); 
                      n = reader.NumberOfPages;
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
    public string GetParameter(SPWeb web, string key)
    {
        string retValue = "";
        try
        {
            SPList list = web.Lists["Parameters"];
            SPQuery spquery = new SPQuery();
            spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + key + "</Value></Eq></Where>";
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
    }
    public void SendErrorEmail(SPWeb web, string Body, string Subject)
    {
        SmtpClient myClient = new SmtpClient(SmtpServer);

        //if (!string.IsNullOrEmpty(GetParameter("FromErrorSmtpUser")))
        //    myClient.Credentials = new System.Net.NetworkCredential(GetParameter("FromErrorSmtpUser"), GetParameter("FromErrorSmtpPWD"));
        //else
        myClient.UseDefaultCredentials = true;
     //   myClient.Credentials = new System.Net.NetworkCredential("darbeirut\\rmzeitouny", "");

        MailMessage msg = new MailMessage();
        msg.From = new MailAddress("PWS_Administrator@dargroup.com", "PWS_Administrator");
        msg.To.Add(GetParameter(web, "ErrorAdmin"));
        //msg.To.Add("Bilal.mustafa@dar.com");
        msg.Subject = Environment.MachineName + " - " + web.Title + " - " + Subject;
        msg.Body = Body;
        msg.IsBodyHtml = true;
        try
        {
            myClient.Send(msg);
        }
        catch { }
    }
    public string SendEmail(SPWeb web, SPListItem item, string to, string cc, string subject, string body, string[] attachement_path1)
    {
        try
        {
            string from = GetParameter(web, "FromRFI");
            string MailFromName = GetParameter(web, "FromRFIName");
            string bcc = GetParameter(web, "AdminEmails");
            string isExternal = GetParameter(web, "isExternal");

            SmtpClient myClient = new SmtpClient(SmtpServer);
            myClient.UseDefaultCredentials = true;
            //myClient.Credentials = new System.Net.NetworkCredential(User, Password);
            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(from, MailFromName);
            msg.BodyEncoding = Encoding.UTF8;

            if (!string.IsNullOrEmpty(to))
                msg.To.Add(to);
            if (!string.IsNullOrEmpty(cc))
                msg.CC.Add(cc);
            if (!string.IsNullOrEmpty(bcc))
                msg.Bcc.Add(bcc);

            body += "<br/>" + "<p><span style='font-family: Verdana;font-size: 13px;color: Red;font-weight: bold;'>Please do not reply to this email as it is sent from an unattended mailbox.</span></p>";

            msg.Subject = subject;
            msg.Body = "";
            msg.Body = body;

            if (isExternal.ToLower() == "yes" || isExternal == "1")
            {
                if (ConfigurationManager.AppSettings["InternalHost"] != null && ConfigurationManager.AppSettings["ExternalHost"] != null)
                {
                    string InternalHost = "";
                    string ExternalHost = "";

                    InternalHost = ConfigurationManager.AppSettings["InternalHost"].ToString();
                    ExternalHost = ConfigurationManager.AppSettings["ExternalHost"].ToString();

                    body = body.Replace(InternalHost, ExternalHost);
                }
            }

            foreach (string s in attachement_path1)
            {
                Attachment att = new Attachment(s);
                msg.Attachments.Add(att);
            }
            msg.IsBodyHtml = true;
            try
            {
                myClient.Send(msg);
                msg.Dispose();
            }
            catch
            {
                //System.Threading.Thread.Sleep(10000);
                //myClient.Send(msg);
            }
        }
        catch
        {
            UpdateLogger(web, item);
        }
        return "Success";
    }
    public string SendEmail(SPWeb web, SPListItem item, MailAddressCollection TO, MailAddressCollection CC, string body, string subject)
    {
        try
        {
            string from = GetParameter(web, "FromRFI");
            string MailFromName = GetParameter(web, "FromRFIName");
            string BCC = GetParameter(web, "AdminEmails");
            string isExternal = GetParameter(web, "isExternal");

            SmtpClient myClient = new SmtpClient(SmtpServer);
            myClient.UseDefaultCredentials = true;

            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(from, MailFromName);

            TO = CleanDuplicates(TO);
            CC = CleanDuplicates(CC);

            if (TO != null && TO.Count > 0)
                foreach (MailAddress em in TO) { msg.To.Add(em); }

            if (CC != null && CC.Count > 0)
                foreach (MailAddress em in CC) { msg.CC.Add(em); }

            if (!string.IsNullOrEmpty(BCC))
                msg.Bcc.Add(BCC);

            body += "<br/>" + "<p><span style='font-family: Verdana;font-size: 13px;color: Red;font-weight: bold;'>Please do not reply to this email as it is sent from an unattended mailbox.</span></p>";

            msg.Subject = subject;
            msg.Body = "";
            msg.Body = body;

            msg.IsBodyHtml = true;
            msg.BodyEncoding = Encoding.UTF8; //Encoding.Unicode;
                                              //    int retry = 0;

            if (isExternal.ToLower() == "yes" || isExternal == "1")
            {
                if (ConfigurationManager.AppSettings["InternalHost"] != null && ConfigurationManager.AppSettings["ExternalHost"] != null)
                {
                    string InternalHost = "";
                    string ExternalHost = "";

                    InternalHost = ConfigurationManager.AppSettings["InternalHost"].ToString();
                    ExternalHost = ConfigurationManager.AppSettings["ExternalHost"].ToString();

                    body = body.Replace(InternalHost, ExternalHost);
                }
            }

            //doUpdate:

            try
            {
                //DateTime start1 = DateTime.Now; //HERE
                myClient.Send(msg);
                msg.Dispose();

                //SendErrorEmail("myClient.Send(msg) = " + DateTime.Now.Subtract(start1).TotalSeconds.ToString() + "<br/>", itemURL);
            }
            catch (SmtpException ex)
            {
                //if (retry < 4)
                //{
                //    while (retry < 4)
                //    {
                //        retry++;
                //        System.Threading.Thread.Sleep(1000);
                //        goto doUpdate;
                //    }
                //}

                UpdateLogger(web, item);
                SendErrorEmail(web, "ItemUrl:<br/>" + web.Url + "<br/><br/>Message: <br/>" + ex.Message + "<br><br>StackTrace: <br/>" + ex.StackTrace, "Error SendEmail");
            }
        }
        catch (System.Exception exc)
        {
            return exc.Message;
        }
        return "Success";
    }
    public string SendEmail(SPWeb web, SPListItem item, MailAddressCollection TO, MailAddressCollection CC, string subject, string body, string HostValue)
    {
        try
        {
            string from = GetParameter(web, "FromRFI");
            string MailFromName = GetParameter(web, "FromRFIName");
            string BCC = GetParameter(web, "AdminEmails");
            string isExternal = GetParameter(web, "isExternal");

            SmtpClient myClient = new SmtpClient(SmtpServer);
            myClient.UseDefaultCredentials = true;

            MailMessage msg = new MailMessage();
            msg.From = new MailAddress(from, MailFromName);

            TO = CleanDuplicates(TO);
            CC = CleanDuplicates(CC);

            if (TO != null && TO.Count > 0)
                foreach (MailAddress em in TO) { msg.To.Add(em); }

            if (CC != null && CC.Count > 0)
                foreach (MailAddress em in CC) { msg.CC.Add(em); }

            if (!string.IsNullOrEmpty(BCC))
                msg.Bcc.Add(BCC);

            msg.Body = "";

            if (!string.IsNullOrEmpty(HostValue))
            {
                string InternalHost = "";
                string ExternalHost = "";

                string href = item.Web.Url + "/" + item.ParentList.RootFolder.Url + "/DispForm.aspx?ID=" + item.ID;

                if (ConfigurationManager.AppSettings["InternalHost"] != null && ConfigurationManager.AppSettings["ExternalHost"] != null)
                {
                    InternalHost = ConfigurationManager.AppSettings["InternalHost"].ToString();
                    ExternalHost = ConfigurationManager.AppSettings["ExternalHost"].ToString();

                    if (!string.IsNullOrEmpty(InternalHost) && !string.IsNullOrEmpty(ExternalHost))
                    {
                        href = href.Replace(InternalHost, ExternalHost);
                    }
                }

                HostValue = HostValue.Replace("[{ViewItem}]", href).Replace("\"", "");

                string Fbody = HostValue + body;

                body = Fbody;
            }

            if (isExternal.ToLower() == "yes" || isExternal == "1")
            {
                if (ConfigurationManager.AppSettings["InternalHost"] != null && ConfigurationManager.AppSettings["ExternalHost"] != null)
                {
                    string InternalHost = "";
                    string ExternalHost = "";

                    InternalHost = ConfigurationManager.AppSettings["InternalHost"].ToString();
                    ExternalHost = ConfigurationManager.AppSettings["ExternalHost"].ToString();

                    body = body.Replace(InternalHost, ExternalHost);
                }
            }

            body += "<br/>" + "<p><span style='font-family: Verdana;font-size: 13px;color: Red;font-weight: bold;'>Please do not reply to this email as it is sent from an unattended mailbox.</span></p>";

            msg.Subject = subject;
            msg.Body = body;
            msg.BodyEncoding = Encoding.UTF8; //Encoding.Unicode;
            msg.IsBodyHtml = true;

            if(DeliverableType == "SLF")
            {
                byte[] ByteArray = Encoding.Unicode.GetBytes(SLF_ITEMS_BODY.ToString());
                var result = Encoding.Unicode.GetPreamble().Concat(ByteArray).ToArray();

                Attachment att = new Attachment(new MemoryStream(), "");
                string pdfName = StrFullRef + "_" + DateTime.Now.ToShortDateString() + "_ITEMS.xls";
                att = new Attachment(new MemoryStream(result), pdfName);
                msg.Attachments.Add(att);
            }
            //int retry = 0;

            //doUpdate:

            try
            {
                myClient.Send(msg);
                msg.Dispose();
            }
            catch (SmtpException ex)
            {
                //if (retry < 4)
                //{
                //    while (retry < 4)
                //    {
                //        retry++;
                //        System.Threading.Thread.Sleep(1000);
                //        goto doUpdate;
                //    }
                //}

                UpdateLogger(web, item);
                SendErrorEmail(objweb, "ItemUrl:<br/>" + item.Url + "<br/><br/>Message: <br/>" + ex.Message + "<br><br>StackTrace: <br/>" + ex.StackTrace, "Error SendEmail");
            }
        }
        catch (System.Exception exc)
        {
            return exc.Message;
        }
        return "Success";
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
    public void UpdateLogger(SPWeb web, SPListItem item)
    {
        if (web.Lists.TryGetList("PWSLogger") != null)
        {
            //SPFieldUserValue Editor = new SPFieldUserValue(web, item["Editor"].ToString());
            string uName = SPContext.Current.Web.CurrentUser.Name; //Editor.User.Name;

            SPList PWSLogger = web.Lists["PWSLogger"];
            SPListItem oSPListItem = PWSLogger.Items.Add();
            oSPListItem["Title"] = web.Url + "/" + item.Url;
            oSPListItem["TransactionType"] = "Email Failure";
            oSPListItem["UserName"] = uName;
            oSPListItem["Date"] = DateTime.Now;
            oSPListItem.Update();
        }
    }
    public bool IsValidEmail(string value)
    {
        Regex r = new Regex(@"^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$");
        if (!string.IsNullOrEmpty(value)) return r.IsMatch(value);
        else return false;
    }
    public bool CheckColumnExistance(SPList GetList, string ColumnName)
    {
        bool Exist = false;
        if (GetList.Fields.ContainsField(ColumnName))
        {
            Exist = true;
        }

        return Exist;
    }
    public void GetStatusFlagFromReviewCodes(SPList InspectionReviewCodesList, string Code, ref bool IsSenttoRE, ref bool IsSenttoDC)
    {
        SPQuery DQquery = new SPQuery();
        DQquery.Query = "<Where>" +
                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + Code + "</Value></Eq>" +
                        "</Where>";

        SPListItemCollection ReviewCodesItems = InspectionReviewCodesList.GetItems(DQquery);

        if (ReviewCodesItems.Count == 1)
        {
            foreach (SPListItem ReviewCodesitem in ReviewCodesItems)
            {
                IsSenttoRE = (ReviewCodesitem["IsSenttoRE"] != null) ? bool.Parse(ReviewCodesitem["IsSenttoRE"].ToString()) : false;
                IsSenttoDC = (ReviewCodesitem["IsSenttoDC"] != null) ? bool.Parse(ReviewCodesitem["IsSenttoDC"].ToString()) : false;
            }
        }
    }
    public string GetEmailsforEInspection(SPWeb web, string mailTo, ref MailAddressCollection mailToCol)
    {
        string GroupName = "";
        int GroupLength = 0;

        try
        {
            if (!string.IsNullOrEmpty(mailTo))
            {
                string[] emailsTo = mailTo.Split(',');

                foreach (string em in emailsTo)
                {
                    if (IsValidEmail(em.Trim()))
                    {
                        mailToCol.Add(em.Trim());
                    }

                    else
                    {
                        GroupName = em.ToString().Trim();

                        if (!string.IsNullOrEmpty(GroupName))
                            GroupLength = GroupName.Length;

                        if (GroupLength <= 250 && !string.IsNullOrEmpty(GroupName))
                        {
                            SPGroup group = web.Groups[GroupName.Trim()];
                            foreach (SPUser usrGroup in group.Users)
                            {
                                if (!string.IsNullOrEmpty(usrGroup.Email))
                                {
                                    if (IsValidEmail(usrGroup.Email))
                                        mailToCol.Add(new MailAddress(usrGroup.Email));
                                }
                            }
                        }
                    }
                }
            }

            return "Success";
        }
        catch (Exception e)
        {
            SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>Error in GetEmailsforEInspection: " + GroupName + " <br/> Message:" + e.ToString(), "Error in GetEmailsforEInspection");
            return web.Url + " Error: " + e.ToString();
        }
    }
    public string CCAssignedToTasks(SPWeb web, SPListItem item, string fieldname)
    {
        string To = "";
        SPFieldUserValueCollection BulkUserfromTradeList = (item[fieldname] != null) ? new SPFieldUserValueCollection(web, item[fieldname].ToString()) : null;
        if (BulkUserfromTradeList != null)
        {
            if (BulkUserfromTradeList.Count != 0)
            {
                for (int i = 0; i < BulkUserfromTradeList.Count; i++)
                {
                    if (!String.IsNullOrEmpty(BulkUserfromTradeList[i].User.Email))
                    {
                        To += BulkUserfromTradeList[i].User.Email + ",";
                    }
                }

                if (!string.IsNullOrEmpty(To))
                    To = To.Remove(To.Length - 1);
            }
        }
        return To;
    }
    public string GetTradeReviewersleadPartEInspection(SPWeb web, string Trade, string CloumName, string DeliverableType)
    {
        ArrayList Fields = new ArrayList();
        Fields.Add(CloumName);  //CDSLead

        string To = "";
        SPList Trades = web.Lists[MainList]; //web.Lists[GetParameter(web, "Inspection-Matrix")];
        SPQuery TradeQ = new SPQuery();
        TradeQ.Query = "<Where><Eq>" +
                            "<FieldRef Name='Title' /><Value Type='Text'>" + Trade + "</Value>" +
                        "</Eq></Where>";
        SPListItemCollection Items = Trades.GetItems(TradeQ);
        foreach (SPListItem item in Items)
        {
            for (int j = 0; j < Fields.Count; j++)
            {
                string fieldname = "";
                fieldname = Fields[j].ToString();

                SPFieldUserValueCollection BulkUserfromTradeList = null;

                if (CheckColumnExistance(Trades, DeliverableType + fieldname))
                    BulkUserfromTradeList = (item[DeliverableType + fieldname] != null) ? new SPFieldUserValueCollection(web, item[DeliverableType + fieldname].ToString()) : null;

                else
                    BulkUserfromTradeList = (item[fieldname] != null) ? new SPFieldUserValueCollection(web, item[fieldname].ToString()) : null;

                if (BulkUserfromTradeList != null)
                {
                    if (BulkUserfromTradeList.Count != 0)
                    {
                        for (int i = 0; i < BulkUserfromTradeList.Count; i++)
                        {
                            if (!String.IsNullOrEmpty(BulkUserfromTradeList[i].User.Email))
                            {
                                if (!To.Contains(BulkUserfromTradeList[i].User.Email))
                                    To += BulkUserfromTradeList[i].User.Email + ",";
                            }
                        }
                    }
                }
            }
        }
        if (!string.IsNullOrEmpty(To))
            To = To.Remove(To.Length - 1);
        return To;
    }
    private string FormatMultipleEmailAddresses(string emailAddresses)
    {
        var delimiters = new[] { ',', ';' };

        var addresses = emailAddresses.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);

        return string.Join(",", addresses);
    }
    public MailAddressCollection CleanDuplicates(MailAddressCollection MailTo)
    {
        MailAddressCollection CopyMailTo = new MailAddressCollection();
        if (MailTo != null && MailTo.Count > 0)
        {
            foreach (MailAddress add in MailTo)
            {
                if (CopyMailTo.Contains(new MailAddress(add.Address, add.DisplayName))) { }
                else
                    CopyMailTo.Add(new MailAddress(add.Address, add.DisplayName));
            }
        }
        return CopyMailTo;
    }
    public bool isValid_In_NotificationList(SPWeb web, string module) //For Multi-Contractor
    {
        bool isValid = false;
        SPList list = web.Lists["Notification"];
        SPQuery spquery = new SPQuery();
        spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + module + "</Value></Eq></Where>";
        SPListItemCollection colitems = list.GetItems(spquery);
        if (colitems.Count > 0)
            isValid = true;

        return isValid;
    }
    public string GetNotificationList(SPWeb web, string module, ref MailAddressCollection mailToCol, ref MailAddressCollection mailCCcol, string ContCode, bool SetContactor_In_To, bool SetContractor_In_CC)
    {
        try
        {
            SPList list = web.Lists["Notification"];
            SPQuery spquery = new SPQuery();
            spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                            module + "</Value></Eq></Where>";
            SPListItemCollection colitems = list.GetItems(spquery);
            if (colitems.Count < 1)
            {
                SendErrorEmail(web, "Project Link: " + web.Url + "<br/>Error in GetNotificationList: Multiple or No records for module " + module + " Count: " + colitems.Count, "Error in GetNotificationList");
                return "";
            }
            SPListItem item = colitems[0];
            string mailTo = (item["To"] != null) ? item["To"].ToString() : "";
            string mailCC = (item["CC"] != null) ? item["CC"].ToString() : "";

            //remove invalid emails 
            string[] emails = mailTo.Split(',');
            foreach (string em in emails)
            {
                if (IsValidEmail(em))
                {
                    mailToCol.Add(em);
                }
            }

            if (item["ToUsers"] != null)
            {
                SPFieldUserValueCollection ToUsers = new SPFieldUserValueCollection(web, item["ToUsers"].ToString());

                foreach (SPFieldUserValue usr in ToUsers)
                {
                    if (usr.User == null) // value is a SPGroup if User is null
                    {
                        SPGroup group = web.Groups[usr.LookupValue];
                        foreach (SPUser usrGroup in group.Users)
                        {
                            if (!string.IsNullOrEmpty(usrGroup.Email))
                            {
                                if (IsValidEmail(usrGroup.Email))
                                    mailToCol.Add(new MailAddress(usrGroup.Email));
                            }
                        }
                    }

                    else // value is SPUser
                    {
                        if (!string.IsNullOrEmpty(usr.User.Email))
                        {
                            if (IsValidEmail(usr.User.Email))
                                mailToCol.Add(new MailAddress(usr.User.Email));
                        }
                    }
                }

            }

            if (SetContactor_In_To)
                GetContractorEmail(web, true, ContCode, ref mailToCol);

            emails = mailCC.Split(',');
            foreach (string em in emails)
            {
                if (IsValidEmail(em))
                    mailCCcol.Add(em);
            }

            if (item["CCUsers"] != null)
            {
                SPFieldUserValueCollection CCUsers = new SPFieldUserValueCollection(web, item["CCUsers"].ToString());

                foreach (SPFieldUserValue usr in CCUsers)
                {
                    if (usr.User == null) // value is a SPGroup if User is null
                    {
                        SPGroup group = web.Groups[usr.LookupValue];
                        foreach (SPUser usrGroup in group.Users)
                        {
                            if (!string.IsNullOrEmpty(usrGroup.Email))
                            {
                                if (IsValidEmail(usrGroup.Email))
                                    mailCCcol.Add(new MailAddress(usrGroup.Email));
                            }
                        }
                    }

                    else // value is SPUser
                    {
                        if (!string.IsNullOrEmpty(usr.User.Email))
                        {
                            if (IsValidEmail(usr.User.Email))
                                mailCCcol.Add(new MailAddress(usr.User.Email));
                        }
                    }
                }
            }

            if (SetContractor_In_CC)
                GetContractorEmail(web, true, ContCode, ref mailCCcol);
            return "Success";
        }
        catch (Exception e)
        {
            SendErrorEmail(web, "Project Link: " + web.Url + "<br/>Error in GetNotificationList module:" + module + " <br/> Message:" + e.ToString(), "Error GetNotificationList");
            return web.Url + " Error: " + e.ToString();
        }
    }
    public string GetDisciplineInfoCc(SPWeb web, SPList list, SPListItem item, string ColumnName)
    {
        string Discipline = "";
        string InfoCc = "";

        if (CheckColumnExistance(list, "Discipline"))
        {
            Discipline = (item["Discipline"] != null) ? item["Discipline"].ToString() : "";

            if (Discipline.Contains("#"))
                Discipline = Discipline.Split('#')[1];

            SPList DisciplineList = web.Lists["Discipline"];

            SPQuery TradeQ = new SPQuery();

            if (DisciplineList.Fields.ContainsField(ColumnName))
            {
                TradeQ.Query = "";
                TradeQ.Query = "<Where>" +
                                    "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + Discipline + "</Value></Eq>" +
                                "</Where>";
                SPListItemCollection Items = DisciplineList.GetItems(TradeQ);
                foreach (SPListItem Discplineitem in Items)
                {
                    SPFieldUserValueCollection BulkUserfromTradeList = (Discplineitem[Discplineitem.Fields.GetFieldByInternalName(ColumnName).Id] != null) ? new SPFieldUserValueCollection(web, Discplineitem[Discplineitem.Fields.GetFieldByInternalName(ColumnName).Id].ToString()) : null;
                    if (BulkUserfromTradeList != null)
                    {
                        if (BulkUserfromTradeList.Count != 0)
                        {
                            for (int i = 0; i < BulkUserfromTradeList.Count; i++)
                            {
                                if (!String.IsNullOrEmpty(BulkUserfromTradeList[i].User.Email))
                                {
                                    if (!InfoCc.Contains(BulkUserfromTradeList[i].User.Email))
                                        InfoCc += BulkUserfromTradeList[i].User.Email + ",";
                                }
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(InfoCc))
                InfoCc = InfoCc.Remove(InfoCc.Length - 1);
        }

        return InfoCc;
    }
    public MailMessage MailMessageCC(string InfoCc, MailAddressCollection MailCC)
    {
        var mailMessage = new MailMessage();

        if (MailCC != null && MailCC.Count > 0)
        {
            foreach (MailAddress em in MailCC) { mailMessage.CC.Add(em); }
        }

        if (!string.IsNullOrEmpty(InfoCc))
            mailMessage.CC.Add(FormatMultipleEmailAddresses(InfoCc));

        return mailMessage;
    }
    public string GetPackageInfoCc(SPWeb web, SPList list, SPListItem item)
    {
        string Package = "";
        if (CheckColumnExistance(list, "Package"))
            Package = (item["Package"] != null) ? item["Package"].ToString() : "";

        if (Package.Contains("#"))
            Package = Package.Split('#')[1];

        string InfoCc = "";

        SPList PackageList = web.Lists["Package"];

        SPQuery TradeQ = new SPQuery();

        string PackageColumnsLists = "ResidentEng;CCUsers;SeniorEngineer";

        List<string> ArrayPackageColumnsLists = PackageColumnsLists.Split(';').ToList();

        string PackageColumn = "";

        TradeQ.Query = "";
        TradeQ.Query = "<Where>" +
                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + Package + "</Value></Eq>" +
                        "</Where>";
        SPListItemCollection Items = PackageList.GetItems(TradeQ);
        foreach (SPListItem Discplineitem in Items)
        {
            for (int j = 0; j < ArrayPackageColumnsLists.Count; j++)
            {
                PackageColumn = ArrayPackageColumnsLists[j].ToString();
                if (PackageList.Fields.ContainsField(PackageColumn))
                {
                    SPFieldUserValueCollection BulkUserfromTradeList = (Discplineitem[PackageColumn] != null) ? new SPFieldUserValueCollection(web, Discplineitem[PackageColumn].ToString()) : null;
                    if (BulkUserfromTradeList != null)
                    {
                        if (BulkUserfromTradeList.Count != 0)
                        {
                            for (int i = 0; i < BulkUserfromTradeList.Count; i++)
                            {
                                if (!String.IsNullOrEmpty(BulkUserfromTradeList[i].User.Email))
                                {
                                    if (!InfoCc.Contains(BulkUserfromTradeList[i].User.Email))
                                        InfoCc += BulkUserfromTradeList[i].User.Email + ",";
                                }
                            }
                        }
                    }
                }
            }

            if (CheckColumnExistance(PackageList, "Cc"))
            {
                string mailCC = (Discplineitem["Cc"] != null) ? Discplineitem["Cc"].ToString() : "";
                InfoCc += mailCC + ",";
            }
        }

        if (!string.IsNullOrEmpty(InfoCc))
            InfoCc = InfoCc.Remove(InfoCc.Length - 1);

        return InfoCc;
    }
    public string GetClientInfoCc(SPWeb web, SPList Notificationlist, string ModuleName)
    {
        string ClientCc = "";
        SPQuery spquery = new SPQuery();
        spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                        ModuleName + "</Value></Eq></Where>";
        SPListItemCollection colitems = Notificationlist.GetItems(spquery);
        if (colitems.Count < 1)
        {
            SendErrorEmail(web, "Project Link: " + web.Url + "<br/>Error in GetNotificationList: Multiple or No records for module " + ModuleName + " Count: " + colitems.Count, "Error GetNotificationList");
            return "";
        }
        SPListItem Notificationitem = colitems[0];
        ClientCc = (Notificationitem["ClientCc"] != null) ? Notificationitem["ClientCc"].ToString() : "";

        return ClientCc;
    }
    public string GetEmailsfromGroupSP(SPWeb web, string mailTo, string mailCC, ref MailAddressCollection mailToCol, ref MailAddressCollection mailCCcol)
    {
        string GroupName = "";
        int GroupLength = 0;

        try
        {
            if (!string.IsNullOrEmpty(mailTo))
            {
                string[] emailsTo = mailTo.Split(',');

                foreach (string em in emailsTo)
                {
                    if (IsValidEmail(em.Trim()))
                    {
                        mailToCol.Add(em.Trim());
                    }

                    else
                    {
                        GroupName = em.ToString().Trim();

                        if (!string.IsNullOrEmpty(GroupName))
                            GroupLength = GroupName.Length;

                        if (GroupLength <= 250 && !string.IsNullOrEmpty(GroupName))
                        {
                            SPGroup group = web.Groups[GroupName.Trim()];
                            foreach (SPUser usrGroup in group.Users)
                            {
                                if (!string.IsNullOrEmpty(usrGroup.Email))
                                {
                                    if (IsValidEmail(usrGroup.Email))
                                        mailToCol.Add(new MailAddress(usrGroup.Email));
                                }
                            }
                        }
                    }
                }
            }

            if (!string.IsNullOrEmpty(mailCC))
            {
                string[] emailsCc = mailCC.Split(',');

                foreach (string em in emailsCc)
                {
                    if (IsValidEmail(em.Trim()))
                    {
                        mailCCcol.Add(em.Trim());
                    }

                    else
                    {
                        GroupName = em.ToString().Trim();

                        if (!string.IsNullOrEmpty(GroupName))
                            GroupLength = GroupName.Length;

                        if (GroupLength <= 250 && !string.IsNullOrEmpty(GroupName))
                        {
                            SPGroup group = web.Groups[GroupName.Trim()];
                            foreach (SPUser usrGroup in group.Users)
                            {
                                if (!string.IsNullOrEmpty(usrGroup.Email))
                                {
                                    if (IsValidEmail(usrGroup.Email))
                                        mailCCcol.Add(new MailAddress(usrGroup.Email));
                                }
                            }
                        }
                    }
                }
            }

            return "Success";
        }
        catch (Exception e)
        {
            SendErrorEmail(web, "Project Link: " + web.Url + "<br/>Error in GetEmailsfromGroupSP: " + GroupName + " <br/> Message:" + e.ToString(), "Error GetEmailsfromGroupSP");
            return web.Url + " Error: " + e.ToString();
        }
    }
    public void CheckExternalExternalHost(SPWeb web, ref string retValue)
    {
        retValue = "";
        try
        {
            SPList list = web.Lists["Parameters"];
            SPQuery spquery = new SPQuery();
            spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>ContractorHostURL</Value></Eq></Where>";
            SPListItemCollection colitems = list.GetItems(spquery);
            if (colitems.Count == 1)
            {
                SPListItem item = colitems[0];
                if (item["Value"] != null && item["Value"].ToString() != "")
                {
                    retValue = item["Value"].ToString().Trim();
                }
            }
            else if (colitems.Count == 0)
            {
                retValue = "";
            }
            else
            {
                SendErrorEmail(web, "itemURL: <br/>" + web.Url + " <br/><br/>" + "Error Message:<br/> Error in GetParameter: Multiple records for Type: ContractorHostURL", "Error CheckExternalExternalHost");
            }
        }
        catch (Exception e)
        {
            if (e.Message.ToString() != "Exception occurred. (Exception from HRESULT: 0x80020009 (DISP_E_EXCEPTION))")
                SendErrorEmail(web, "itemURL: <br/>" + web.Url + "<br/><br/>" +
                           "Error Message:<br/> Error in GetParameter Type: ContractorHostURL<br/><br/> " +
                           "Message: <br/>" + e.Message + "<br/><br/>" +
                           "StackTrace: <br/>" + e.StackTrace, "Error CheckExternalExternalHost");
        }
    }
    public void GetContractorEmail(SPWeb web, bool Get_ContractorEmail_From_List, string ContractorCode, ref MailAddressCollection GetUserEmails)
    {
        if (Get_ContractorEmail_From_List)
        {
            string ContractorList = GetParameter(web, "ContractorCode");
            SPList Contractor = web.Lists[ContractorList];
            SPQuery Contquery = new SPQuery();
            Contquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + ContractorCode + "</Value></Eq></Where>";
            SPListItemCollection Contqueryitems = Contractor.GetItems(Contquery);

            if (Contqueryitems.Count > 0)
            {
                SPListItem Contitem = Contqueryitems[0];

                string ContEmail = (Contitem["Email_User"] != null) ? Contitem["Email_User"].ToString().Trim().Replace(" ", "") : "";
                if (!string.IsNullOrEmpty(ContEmail))
                {
                    string[] emails = ContEmail.Split(',');
                    foreach (string em in emails)
                    {
                        if (IsValidEmail(em))
                            GetUserEmails.Add(new MailAddress(em));

                        else
                        {
                            int GroupLength = 0;
                            string GroupName = em.ToString().Trim();

                            if (!string.IsNullOrEmpty(GroupName))
                                GroupLength = GroupName.Length;

                            if (GroupLength <= 250 && !string.IsNullOrEmpty(GroupName))
                            {
                                SPGroup group = web.Groups[GroupName.Trim()];
                                foreach (SPUser usrGroup in group.Users)
                                {
                                    if (!string.IsNullOrEmpty(usrGroup.Email))
                                    {
                                        if (IsValidEmail(usrGroup.Email))
                                            GetUserEmails.Add(new MailAddress(usrGroup.Email));
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    public bool GetNotificationModuleName(SPWeb web, string key)
    {
        bool retValue = false;

        SPList list = web.Lists["Notification"];
        SPQuery spquery = new SPQuery();
        spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                        key + "</Value></Eq></Where>";
        SPListItemCollection colitems = list.GetItems(spquery);
        if (colitems.Count == 1)
        {
            retValue = true;
        }

        return retValue;
    }

    public void EXTRACT_MAJORTYPE_LIST_VALUES(SPWeb web, string DeliverableType)
    {
        SPList MajorTypes = web.Lists["MajorTypes"];

        if (!CheckColumnExistance(MajorTypes, "isSiteModule")) return;

        SPQuery spquery = new SPQuery();
        spquery.Query = "<Where>" +
                           "<And>" +
                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + DeliverableType + "</Value></Eq>" +
                            "<Eq><FieldRef Name='isSiteModule' /><Value Type='Boolean'>1</Value></Eq>" +
                           "</And>" +
                        "</Where>";
        SPListItemCollection colitems = MajorTypes.GetItems(spquery);
        if (colitems.Count > 0)
        {
            SPListItem ModuleItem = colitems[0];
            MainList = (ModuleItem["MainList"] != null) ? ModuleItem["MainList"].ToString() : "";
            LeadTaskList = (ModuleItem["LeadList"] != null) ? ModuleItem["LeadList"].ToString() : "";
            PartTaskList = (ModuleItem["PartList"] != null) ? ModuleItem["PartList"].ToString() : "";
            FormCodes = (ModuleItem["Code"] != null) ? ModuleItem["Code"].ToString() : "";
            this.Category = (ModuleItem["Category"] != null) ? ModuleItem["Category"].ToString() : "";
            FormName = (ModuleItem["DigitalFormName"] != null) ? ModuleItem["DigitalFormName"].ToString() : "";
            //isCompile = (ModuleItem["isCompile"] != null) ? bool.Parse(ModuleItem["isCompile"].ToString()) : false;
            _IssuedItem_Email = (ModuleItem["IssuedEmail"] != null) ? ModuleItem["IssuedEmail"].ToString() : "";
            _LeadTaskClose_Email = (ModuleItem["LeadTaskClose_Email"] != null) ? ModuleItem["LeadTaskClose_Email"].ToString() : "";
        }
    }

    public int GET_SLF_OPENITEMS_COUNT(SPWeb web, string Reference)
    {
        SPList list = web.Lists["Snag Items"];
        SPQuery spquery = new SPQuery();
        spquery.Query = "<Where><And>" +
                            "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + Reference + "</Value></Eq>" +
                            "<Eq><FieldRef Name='Status' /><Value Type='Text'>Open</Value></Eq>" +
                          "</And></Where>";
        SPListItemCollection colitems = list.GetItems(spquery);
        return colitems.Count;
    }

    public static string GetFileUrl(string hrefElement)
    {
        string fileUrl = "";
        Regex regex = new Regex("href\\s*=\\s*(?:\"(?<1>[^\"]*)\"|(?<1>\\S+))", RegexOptions.IgnoreCase);
        Match match;

        for (match = regex.Match(hrefElement); match.Success; match = match.NextMatch())
        {
            foreach (Group linkurl in match.Groups)
            {
                fileUrl = linkurl.Value.Replace("href='", "").Replace("href=", "").Replace("\"", "");
                fileUrl.Trim('\'');
            }
        }
        return fileUrl;
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
        //if (Referrer != "")
        //    Response.Write("<script>window.location = '" + Referrer + "';</script>");
        //else
            Response.Write("<script>window.location = '" + url + "';</script>");
    }

    protected void btnSubmitCloseIR_Click(object sender, EventArgs e)
    {
        SPWeb objweb = null;
        DateTime Start = DateTime.Now;
        string TimeLog = "";

        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    objweb = site.OpenWeb();
                }
            });
            objweb.AllowUnsafeUpdates = true;           

            if (!IgnorePDFFile)
            {
                if (LeadItem.Attachments.Count == 0)
                {
                    lblmsg.Text = "No Compiled Document attachements found on Lead task";
                    return;
                }
                else if (LeadItem.Attachments.Count > 1)
                {
                    lblmsg.Text = "Multiple attachements on Lead task found, only one final attachment required";
                    return;
                }
            }

            if (string.IsNullOrEmpty(ddlcode.SelectedValue)) //LeadItem["Code"] == null)
            {
                lblmsg.Text = "Code is empty ";
                return;
            }
            else if (DeliverableType == "SLF" && isTaskOpen)
            {
                lblmsg.Text = "Inspectors Tasks must be closed before compiling the pdf";
                return;
            }
            else
            {
                Compile_Close(objweb, LeadItem);                
            }

            int total = (int)DateTime.Now.Subtract(Start).TotalSeconds;
            if (total > 25)
            {
                TimeLog += "TOTAL TIME = " + total;
                SendErrorEmail(objweb, "btnSubmitCloseIR_Click = " + TimeLog, "Lead Submit");
            }

            //MessageAndRedirect2("Submitted Successfully, please click Ok", SPContext.Current.Web.Url + "/Lists/LeadAction"); //+ Referrer
            Response.Write("<script>window.location = '" + SPContext.Current.Web.Url + "/Lists/LeadAction" + "';</script>");
            return;

            //Response.Redirect(SPContext.Current.Web.Url + "/Lists/LeadAction");
        }
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator";
            SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>btnSubmitCloseIR_Click Error: Item Ref No: " + lblIRNo.Text + "<br/>" + ex.ToString(), ErrorSubject);
        }
        finally { if (objweb != null) objweb.Dispose(); }
    }

    protected void btnSave_Click(object sender, EventArgs e)
    {
        SPWeb objweb = null;

        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    objweb = site.OpenWeb();
                }
            });
            objweb.AllowUnsafeUpdates = true;

            lblmsg.Text = "";
            string _Code = ddlcode.SelectedValue;
            if (_Code == "Select Code")
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
            LeadItem["Code"] = _Code;
            LeadItem.Update();  //SystemUpdate(false);

            lblmsg.Text = "Saved successfully.";
            SetButtons();
        }
        catch (Exception ex) {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>Page Load Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> Ref:" + lblIRNo.Text + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        finally { if (objweb != null) objweb.Dispose(); }
    }

    protected void btnUpdateAREMergePdf_Click(object sender, EventArgs e)
    {
        SPWeb objweb = null;
        try
        {
            SPSecurity.RunWithElevatedPrivileges(delegate ()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID)) //
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

            if(DeliverableType == "SLF")
            {
                if (isTaskOpen)
                {
                    lblmsg.Text = "Inspectors Tasks must be closed before compiling the pdf";
                    return;
                }
            }

            foreach (GridViewRow gritemcheck in grdWorkflowtasks.Rows)
            {
                Label lblIDcheck = (Label)gritemcheck.FindControl("lblID");
                Label lblWTStatus = (Label)gritemcheck.FindControl("lblWTStatus");
                Label lblAssignedTo = (Label)gritemcheck.FindControl("lblAssignedTo");
                //Label lblSeniorEngineer = (Label)gritemcheck.FindControl("lblSeniorEngineer");
                CheckBoxList chkBxListcheck = (CheckBoxList)gritemcheck.FindControl("Chkattachments");

                if (chkBxListcheck != null)
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

                    foreach (System.Web.UI.WebControls.ListItem itemcheck in chkBxListcheck.Items)
                    {
                        if (itemcheck.Selected)
                        {
                            string UpperCaseFilename = itemcheck.Value.ToUpper();
                            intotherfilescount++;

                            if (UpperCaseFilename.Contains(".PDF")) // (itemcheck.Value.ToLower().Contains("engineer")) &&
                            {
                                intEngineersComments++;
                            }
                            else if (UpperCaseFilename.Contains(".JPG") || UpperCaseFilename.ToLower() == ".JPEG" || UpperCaseFilename.Contains(".PNG")) // (itemcheck.Value.ToLower().Contains("engineer")) &&
                            {
                                intimagecount++;
                            }
                        }
                    }
                }
            }
            #endregion

            #region Code for Updating Lead Task with Pdf attachment
            Label2.Text = "";
            List<String> ListPdffiles = new System.Collections.Generic.List<String>();
            if (RIWItem != null && RIWItem.Attachments.Count > 0)
                ListPdffiles.Add(RIWItem.Attachments.UrlPrefix + RIWItem.Attachments[0]);

            #region Bind Images and get commented pdf attachments
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
                CheckBoxList chkBxList = (CheckBoxList)gritem.FindControl("Chkattachments");
                if (chkBxList != null)
                {
                    foreach (System.Web.UI.WebControls.ListItem item in chkBxList.Items)
                    {
                        if (item.Selected)
                        {
                            string File_Ext = Path.GetExtension(item.Value);
                            if (!string.IsNullOrEmpty(File_Ext))
                                File_Ext = File_Ext.ToLower();
                            else continue;

                            string fileUrl = GetFileUrl(item.Text);
                            if (string.IsNullOrEmpty(fileUrl))
                                continue;

                            if (File_Ext == ".pdf")
                              ListPdffiles.Add(fileUrl);
                            
                            #region Bind Images attachments
                            if (File_Ext == ".jpg" || File_Ext == ".jpeg" || File_Ext == ".png")
                            {
                                doc.NewPage();
                                SPFile _itemAttachment = objweb.GetFile(fileUrl);

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
            byte[] MergedAtt = null;

            if (ListPdffiles != null && ListPdffiles.Count > 0)
            {
                MergePDFs(ListPdffiles);

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
                        for (int page = 0; page < n;)
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
            }
            #endregion

            #region CoverPage and comments location
            string CommentsLocation = rblist_CommentsLoc.SelectedValue;
            if (rblist_CommentsLoc.SelectedItem == null)
            {
                lblmsg.Text = "Please choose code and comments placement options.";
                return;
            }

            string EmailBody = "", EmailSubject = "";
            if (DeliverableType == "IR")
            {
                string BldgType = (RIWItem["InspType"] != null) ? RIWItem["InspType"].ToString().Split('#')[1] : "";
                if (BldgType == "Buildings")
                    FormName = "RIW_BLDG";
                else FormName = "OTHER_RIW";
            }

            string SLF_ITEMS_BODY = "";
            GenericEmailTemplate(objweb, FormName, RIWItem, ref EmailBody, ref EmailSubject); //MAIN FOR SLF AND OTHER MAIN MODULES
            if (DeliverableType == "SLF")
            {
                isNewRow = false;
                GenericEmailTemplate(objweb, "SLF_FORM_ITEMS", RIWItem, ref SLF_ITEMS_BODY, ref EmailSubject); // ITEMS FOR SLF
            }

            string FormBody = EmailBody + SLF_ITEMS_BODY; //htmlforPdfFile(objweb, LeadItem);
            byte[] CRSFrom_PDFByte = PDFConvert(FormBody);
            FinalContent = MergeBytePDFs(CRSFrom_PDFByte, MergedAtt);
            #endregion

            if (intimagecount > 0) FinalContent = MergeBytePDFs(FinalContent, BindImagesFileStream.GetBuffer());

            if (ckb_ApplyWaterMark.Checked) FinalContent = ApplyWatermark(FinalContent);
            #endregion

            #region Attach Response To Lead task and update
            string attachName = StrFullRef + "-Commented.pdf";
            //Delete old Attachement
            if (LeadItem.Attachments.Count > 0)
            {
                try { LeadItem.Attachments.DeleteNow(attachName); } catch { }
            }

            LeadItem["PartComments"] = PartComments.Text;
            LeadItem["Comment"] = CommentBox.Text;
            LeadItem["Code"] = ddlcode.SelectedValue;

            LeadItem.Attachments.Add(attachName, FinalContent); //File.ReadAllBytes(strpath3)
            objweb.AllowUnsafeUpdates = true;
            LeadItem.Update();// SystemUpdate(false);

            string attachmentUrl = SPContext.Current.Web.Url + "/" + LeadItem.ParentList.RootFolder.Url + "/Attachments/" + LeadItem.ID + "/" + attachName;

            HyperLinkLeadAttach.NavigateUrl = attachmentUrl;//LeadItem.Attachments.UrlPrefix + attachName;
            HyperLinkLeadAttach.Text = attachName;// "Download Compiled RIW Document";

            embedpdfViewer.Visible = true;
            embedpdfViewer.Attributes.Add("src", HyperLinkLeadAttach.NavigateUrl + "#page=1&zoom=150");

            lblmsg.Text = "Success. Please review the compiled response previewed above before submiting and closing the " + Category + ".";

            //Response.Write("<script>window.open(\"" + LeadItem.Attachments.UrlPrefix + attachName + "\");</script>");
            //Response.Clear();
            //Response.ContentType = "application/pdf";
            //Response.AppendHeader("Content-Disposition", "inline; filename=" + attachName);
            //Response.TransmitFile(LeadItem.Attachments.UrlPrefix + attachName);
            //Response.End(); 
            #endregion

            SubmitCloseIR.Visible = true;
            ddlcode.Enabled = false;
            #endregion
        }
        catch (Exception ex)
        {
            lblmsg.Text = _ColumnName + ", An error occurred. Please contact your PWS system administrator" + "<br/>" + ex.ToString();
            SendErrorEmail(objweb, "Project Link: " + objweb.Url + "<br/>btnUpdateAREMergePdf_Click Error: <br/><br/> User " + SPContext.Current.Web.CurrentUser.Name + "<br/><br/> Reference:" + lblIRNo.Text + "<br/> WebUrl: " + objweb.Url + "<br/><br/>" + ex.ToString(), ErrorSubject);
        }
        finally { objweb.Dispose(); }
    }

    protected void btnCancel_Click(object sender, EventArgs e)
    {
        //if (Referrer != "")
           // Response.Write("<script>window.location = '" + Referrer + "';</script>");
        //else
            Response.Write("<script>window.location = '" + SPContext.Current.Web.Url + "/Lists/LeadAction';</script>");
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
        //EmailQuery.ViewAttributes = "Scope=\"Recursive\"";
        SPListItemCollection items = Emails.GetItems(EmailQuery);
        SPListItem item = items[0];
        EmailSubject = (item["EmailSubject"] != null) ? item["EmailSubject"].ToString() : "";
        EmailBody = (item["Body"] != null) ? item["Body"].ToString() : "";
        string ListName = (item["ListName"] != null) ? item["ListName"].ToString() : "";
        Dictionary<string, string> EmailData = new Dictionary<string, string>();
        EmailData.Add("Subject", EmailSubject);
        EmailData.Add("Body", EmailBody);

        //string Deliv = GetParameter(web, "UploadDeliverables");
        #endregion

        SPList List = null;
        if (!string.IsNullOrEmpty(ListName))
            List = web.Lists[ListName];

        if (List == null)
            List = Listitem.ParentList;


        //Listitem = List.GetItemById(10);

        List<SPListItem> ListItems = new List<SPListItem>();
        if (ListName == "Snag Items")
        {
            SPQuery _query = new SPQuery();
            _query.Query = "<Where>" +
                             "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + StrFullRef + "</Value></Eq>" +
                           "</Where>";
            SPListItemCollection Selecteditems = List.GetItems(_query);
            foreach (SPListItem _item in Selecteditems)
            {
                ListItems.Add(_item);
            }
        }
        else
        {
            if (!ListItems.Contains(Listitem))
                ListItems.Add(Listitem);
        }

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
                        if (DeliverableType == "SLF" && ColumnName == "FullRef")
                            ColumnName = "Reference";
                        _ColumnName = ColumnName;

                        //if (!List.Fields.ContainsField(_ColumnName))
                            //Response.Write(_ColumnName + "<br/>");
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
                            if (ColumnListName == "Snag List")
                            {
                                GenerateQuery.Query = "<Where>" +
                                                        "<Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq>" +
                                                      "</Where>";
                            }
                            #endregion

                            SPList _List = web.Lists[ColumnListName];
                            SPListItemCollection Selecteditems = _List.GetItems(GenerateQuery);
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
                            string Display_Name = _List.Fields.GetField(ColumnName).Title;
                            SPField field = _List.Fields[Display_Name];
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
                                        value = GetParameter(web, "ProjectName");
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
                                value = GetParameter(web, "ProjectName");
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
                                string isMultiContracotr = GetParameter(web, "isMultiContracotr");
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
                            value = GetParameter(web, "ProjectTitle");
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
                                        foreach (GridViewRow gritem in grdWorkflowtasks.Rows)
                                        {
                                            Label lblID = (Label)gritem.FindControl("lblID");
                                            CheckBox IncludePartCommentCheckBox = (CheckBox)gritem.FindControl("IncludePartCommentCheckBox");

                                            if (listItem.ID.ToString() == lblID.Text && IncludePartCommentCheckBox.Checked)
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
                                        }
                                    }

                                    value += "</table>";
                                }
                            }
                            EmailBody = EmailBody.Replace(ColumnInBrackets, value);
                            continue;
                            #endregion
                        }
                        else if (ColumnName == "MIRInspectionReport")
                        {
                            #region MIR INSPECTION REPORT COLUMN
                            value = "";

                            SPList list = web.Lists.TryGetList(PartTaskList);
                            if (list != null)
                            {
                                SPQuery listQuery = new SPQuery();
                                listQuery.Query = "<Where><And>" +
                                                   "<Eq><FieldRef Name='Reference' /><Value Type='Text'>" + Listitem["FullRef"].ToString() + "</Value></Eq>" +
                                                   "<Eq><FieldRef Name='Role' /><Value Type='Text'>TeamLeader</Value></Eq>" +
                                                  "</And></Where>";
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
                        else if (ColumnName == "REName")
                        {
                            value = LeadName;
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

                            string TentativeListName = GetParameter(web, "TentativeLOD");

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
                                value = GetParameter(web, "ProjectName");
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
                            else if (field.Type == SPFieldType.URL)
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
                            else if (field.Type == SPFieldType.DateTime)
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
                            else if (field.Type == SPFieldType.Lookup)
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
                            else if (field.Type == SPFieldType.User)
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
                            else if (field.Type == SPFieldType.MultiChoice)
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
                            else if (field.Type == SPFieldType.Attachments)
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
                            else if (field.Type == SPFieldType.Calculated)
                            {
                                #region SPFieldType.Calculated
                                value = (RLODitem[ColumnName] != null) ? RLODitem[ColumnName].ToString() : "";
                                if (value.Contains('#'))
                                    value = value.Split('#')[1].ToString();

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

    public void Compile_Close(SPWeb web, SPListItem LeadItem)
    {
        //SPWeb web = null;
        //string MainList = "Inspection Request", LeadTaskList = "Inspection Lead Tasks", PartTaskList = "Inspection Part Tasks";
        try
        {
            //SPSecurity.RunWithElevatedPrivileges(delegate ()
            //{
            //    using (SPSite site = new SPSite(SPContext.Current.Web))
            //    { web = site.OpenWeb(); }
            //});

            //web.AllowUnsafeUpdates = true;

            SPList mainList = web.Lists[MainList];
            SPList leadlist = web.Lists[LeadTaskList];
            SPList Partlist = web.Lists[PartTaskList];
            string Trade = (LeadItem["Trade"] != null) ? LeadItem["Trade"].ToString() : "";

            if (isCompile)
            {
                #region AUTO CLOSE OPEN PART ACTION ITEMS
                SPQuery spquery = new SPQuery();
                spquery.Query = "<Where><And><Eq>" +
                                    "<FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq><Eq>" +
                                    "<FieldRef Name='Status' /><Value Type='Choice'>Open</Value>" +
                                "</Eq></And></Where>";
                spquery.ViewFields = "<FieldRef Name='AssignedTo' />" +
                                        "<FieldRef Name='Reviewed' />" +
                                        "<FieldRef Name='Status' />" +
                                        "<FieldRef Name='Trade' />";
                SPListItemCollection PartItems = Partlist.GetItems(spquery);

                if (PartItems.Count > 0)
                {
                    foreach (SPListItem PartItem in PartItems)
                    {
                        #region SET PART ACTION TASK CLOSED
                        SPFieldUserValueCollection AssignedTo = (PartItem["AssignedTo"] != null) ? new SPFieldUserValueCollection(objweb, PartItem["AssignedTo"].ToString()) : null;
                        PartItem["Reviewed"] = true;
                        PartItem["Status"] = "Completed";
                        PartItem.Update();
                        #endregion

                        #region MESSAGE BODY

                        string LeadBody = "";

                        LeadBody += "<p style='font-family: Verdana; font-size: 13px; font-weight: bold;'>Task Assigned for you for " +
                                    "<a href ='" + PartItem.Web.Url + "/" + PartItem.ParentList.RootFolder.Url + "/DispForm.aspx?ID=" + PartItem.ID + "'>" + StrFullRef + "</a>" +
                                    " has been Closed Automatically, as Lead has Closed his Task." +
                                    "<br />";
                        #endregion

                        #region SEND EMAIL TO PART ACTION

                        string To = "";

                        for (int i = 0; i < AssignedTo.Count; i++)
                        {
                            To += AssignedTo[i].User.Email + ",";
                        }

                        if (!string.IsNullOrEmpty(To))
                            To = To.Remove(To.Length - 1);

                        string CurrentUser = SPContext.Current.Web.CurrentUser.Email;
                        if (string.IsNullOrEmpty(CurrentUser))
                            CurrentUser = "";

                        string PartTrade = PartItem["Trade"] != null ? PartItem["Trade"].ToString() : "SLF Coordinator";
                        string subject = GetParameter(objweb, "ProjectName") + " - " + StrFullRef + " - " + PartTrade + " Part Action Task Closed Automatically";
                        SendEmail(objweb, PartItem, To, CurrentUser, subject, LeadBody, new string[] { });
                        #endregion
                    }
                }
                #endregion

                #region SET LEAD RESPONSE INTO MAIN LIST ITEM
                SPFieldUrlValue ReviewLink = new SPFieldUrlValue();
                ArrayList PartDepartments = new ArrayList();
                string RefColumnName = "FullRef";
                if (DeliverableType == "SLF")
                    RefColumnName = "Reference";

                spquery = new SPQuery();
                spquery.Query = "<Where>" +
                                    "<Eq><FieldRef Name='" + RefColumnName + "' /><Value Type='Text'>" + StrFullRef + "</Value></Eq>" +
                                "</Where>";
                spquery.ViewAttributes = "Scope=\"Recursive\"";

                SPListItemCollection Items = mainList.GetItems(spquery);

                if (Items.Count > 0)
                {
                    SPListItem item = Items[0];
                    string DQItemUrl = item.Url.ToString();

                    #region GET PART TRADES               
                    string LeadAction = "";
                    if (CheckColumnExistance(objlist, "LeadAction"))
                    {
                        SPFieldLookupValue LeadTrade = (item["LeadAction"] != null) ? new SPFieldLookupValue(item["LeadAction"].ToString()) : null;
                        if (LeadTrade != null)
                        {
                            LeadAction = LeadTrade.LookupValue;
                            PartDepartments.Add(LeadAction);
                        }
                    }

                    if (CheckColumnExistance(objlist, "PartTrades"))
                    {
                        string PartTrades = (item["PartTrades"] != null) ? item["PartTrades"].ToString() : "";
                        if (!string.IsNullOrEmpty(PartTrades))
                        {
                            string[] SplitPart = PartTrades.Split(',');
                            for (int i = 0; i < SplitPart.Count(); i++)
                            {
                                if (SplitPart[i].ToString() != LeadAction)
                                    PartDepartments.Add(SplitPart[i].ToString());
                            }
                        }
                    }
                    #endregion

                    #region EXTRACT VALUES
                    string Code = ddlcode.SelectedValue;
                    MailAddressCollection MailInspectionOutCO = new MailAddressCollection();
                    MailAddressCollection MailInspectionOutDC = new MailAddressCollection();

                    bool isInspectionOutCO = false;
                    bool isInspectionOutDC = false;
                    bool isToContractorDirectly = false;
                    bool IsSenttoRE = true;
                    bool IsSenttoDC = true;

                    if (!string.IsNullOrEmpty(InspectionOutCO))
                    {
                        isInspectionOutCO = true;
                        GetEmailsforEInspection(objweb, InspectionOutCO, ref MailInspectionOutCO);
                    }
                    if (!string.IsNullOrEmpty(InspectionOutDC))
                    {
                        isInspectionOutDC = true;
                        GetEmailsforEInspection(objweb, InspectionOutDC, ref MailInspectionOutDC);
                    }
                    #endregion

                    #region UPDATE METADATA
                    item["FinalResponse"] = CommentBox.Text;
                    item["Response"] = LeadItem["PartComments"];

                    if (CheckColumnExistance(mainList, "AttachFiles"))
                        item["AttachFiles"] = "";

                    if (string.IsNullOrEmpty(LeadName))
                        LeadName = SPContext.Current.Web.CurrentUser.Name;
                    item["HandledBy"] = LeadName;

                    if (!string.IsNullOrEmpty(InspectionReviewCodes))
                    {
                        SPList InspectionReviewCodesList = objweb.Lists[InspectionReviewCodes];

                        if (!string.IsNullOrEmpty(Code))
                            GetStatusFlagFromReviewCodes(InspectionReviewCodesList, Code, ref IsSenttoRE, ref IsSenttoDC);
                    }

                    if (isInspectionOutCO && IsSenttoRE)
                    {
                        item["Status"] = "Returned to CO";
                        item["SentToPMCDate"] = DateTime.Now.ToString("MMM dd, yyyy");
                    }

                    else if (isInspectionOutDC && IsSenttoDC)
                    {
                        item["Status"] = "Returned to Site";
                        item["SentToSiteDate"] = DateTime.Now.ToString("MMM dd, yyyy");
                    }

                    else
                    {
                        item["Status"] = "Issued to Contractor";
                        item["SentToContractorDate"] = DateTime.Now.ToString("MMM dd, yyyy");
                        item["State"] = "Answered";
                        item["SentToContractor"] = false;
                        isToContractorDirectly = true;
                    }

                    if (!string.IsNullOrEmpty(Code))
                        item["Code"] = Code;

                    if (CheckColumnExistance(mainList, "Editedby"))
                        item["Editedby"] = LeadName;

                    if (CheckColumnExistance(mainList, "ReviewedURL"))
                    {
                        if (LeadItem.Attachments.Count > 0)
                        {
                            string CurrentdestinUrl = LeadItem.Attachments.UrlPrefix + LeadItem.Attachments[0];
                            SPFile file = objweb.GetFile(CurrentdestinUrl);

                            string destinUrl = item.Attachments.UrlPrefix + file.Name.Trim();

                            ReviewLink.Url = destinUrl;
                            ReviewLink.Description = file.Name.Trim();
                            item["ReviewedURL"] = ReviewLink;
                        }
                    }
                    item.Update();

                    for (int i = 0; i < LeadItem.Attachments.Count; i++)
                    {
                        string fURL = LeadItem.Attachments.UrlPrefix + LeadItem.Attachments[i];
                        SPFile f = objweb.GetFile(fURL);
                        string fName = f.Name;// Trade.Replace("&", "") + "_" + f.Name;
                        bool insert = false;
                        for (int j = 0; j < item.Attachments.Count; j++)
                        {
                            string DQFileName = item.Attachments[j];
                            bool x = item.Attachments[j].Equals(fName);
                            if ((item.Attachments[j] != null) && (item.Attachments[j].Equals(fName)))
                                insert = true;
                        }

                        if (insert == false)
                        {
                            item.Attachments.AddNow(fName, f.OpenBinary());
                            LeadItem.Attachments.DeleteNow(f.Name);
                        }
                        else
                        {
                            item.Attachments.DeleteNow(fName);
                            item.Attachments.AddNow(fName, f.OpenBinary());
                        }
                    }
                    #endregion

                    #region GENERIC EMAIL (BODY, SUBJECT)
                    string EmaSubject = "",
                            Body = "";

                    spquery = new SPQuery();
                    spquery.Query = "<Where>" +
                                        "<Eq><FieldRef Name='" + RefColumnName + "' /><Value Type='Text'>" + StrFullRef + "</Value></Eq>" +
                                    "</Where>";
                    spquery.ViewAttributes = "Scope=\"Recursive\"";
                    SPListItemCollection listitems = mainList.GetItems(spquery);

                    if (listitems.Count > 0)
                    {
                        SPListItem listitem = listitems[0];
                        if (isToContractorDirectly)
                        {
                            GenericEmailTemplate(objweb, _IssuedItem_Email, listitem, ref Body, ref EmaSubject);
                            if (DeliverableType == "SLF")
                            {
                                isNewRow = false;
                                string DummySubject = "";
                                GenericEmailTemplate(objweb, "SLF_FORM_ITEMS", listitem, ref SLF_ITEMS_BODY, ref DummySubject);
                            }
                        }
                        else
                            GenericEmailTemplate(objweb, _LeadTaskClose_Email, listitem, ref Body, ref EmaSubject);
                    }
                    #endregion

                    #region SEND LEAD ACTION EMAIL
                    string To = "";

                    To = CCAssignedToTasks(objweb, LeadItem, "AssignedTo");

                    var mailMessage = new MailMessage();
                    MailAddressCollection mailCCcol = new MailAddressCollection();

                    #region MIR TRADES
                    //string DeliverableType = "MIR";
                    for (int j = 0; j < PartDepartments.Count; j++)
                    {
                        string Dept2 = PartDepartments[j].ToString();
                        string PartUsers = GetTradeReviewersleadPartEInspection(objweb, Dept2, "CC", DeliverableType);

                        if (!string.IsNullOrEmpty(PartUsers))
                            To += "," + PartUsers;

                        PartUsers = GetTradeReviewersleadPartEInspection(objweb, Dept2, "RE", DeliverableType);
                        if (!string.IsNullOrEmpty(PartUsers))
                            To += "," + PartUsers;

                        PartUsers = GetTradeReviewersleadPartEInspection(objweb, Dept2, "TeamLeader", DeliverableType);
                        if (!string.IsNullOrEmpty(PartUsers))
                            To += "," + PartUsers;

                        PartUsers = GetTradeReviewersleadPartEInspection(objweb, Dept2, "Inspectors", DeliverableType);
                        if (!string.IsNullOrEmpty(PartUsers))
                            To += "," + PartUsers;
                    }
                    #endregion

                    if (!string.IsNullOrEmpty(To))
                    {
                        mailMessage.CC.Add(FormatMultipleEmailAddresses(To));
                    }

                    string subject = GetParameter(objweb, "ProjectName") + " - (" + Trade + ") - " + EmaSubject;
                    if (isInspectionOutCO && IsSenttoRE)
                        SendEmail(objweb, item, MailInspectionOutCO, mailMessage.CC, Body, subject);

                    else if (isInspectionOutDC && IsSenttoDC)
                        SendEmail(objweb, item, MailInspectionOutDC, mailMessage.CC, Body, subject);

                    else
                    {
                        #region CHECK IF MULTI-CONTRACTOR IS ENABLED
                        string isMultiContracotr = GetParameter(objweb, "isMultiContracotr");
                        string _Notification_Name = "IR Initiated";

                        if (DeliverableType == "SCR")
                            _Notification_Name = "RFI Initiated";

                        string ContCode = "";
                        bool isValid = false;
                        if (isMultiContracotr.ToLower() == "yes" || isMultiContracotr == "1")
                        {
                            string tmp = DQItemUrl.Substring(DQItemUrl.ToLower().IndexOf("/lists/") + 7); //Strip /Lists/ and preceding chars
                            string Listname = tmp.Substring(0, tmp.IndexOf("/"));
                            tmp = tmp.Substring(tmp.IndexOf("/") + 1); //Strip "Listname/"
                            if (tmp.Contains("/"))
                                ContCode = tmp.Substring(0, tmp.IndexOf("/"));
                            else
                                isMultiContracotr = "0";

                            isValid = isValid_In_NotificationList(objweb, ContCode + "_" + _Notification_Name);
                            if (!isValid)
                            {
                                isValid = isValid_In_NotificationList(objweb, DeliverableType + " Initiated");
                                if (isValid)
                                    _Notification_Name = DeliverableType + " Initiated";
                            }
                            else _Notification_Name = ContCode + "_" + _Notification_Name;
                        }
                        else
                        {
                            isValid = isValid_In_NotificationList(objweb, DeliverableType + " Initiated");
                            if (isValid)
                                _Notification_Name = DeliverableType + " Initiated";
                        }
                        #endregion

                        MailAddressCollection MailTo = new MailAddressCollection();
                        MailAddressCollection MailCC = new MailAddressCollection();

                        //GetNotificationList(web, ModuleName, ref MailTo, ref MailCC, "", false, false);

                        if (isMultiContracotr.ToLower() == "yes" || isMultiContracotr == "1")
                            GetNotificationList(objweb, _Notification_Name, ref MailTo, ref MailCC, ContCode, false, true);
                        else GetNotificationList(objweb, _Notification_Name, ref MailTo, ref MailCC, "", false, false);

                        if (MailTo != null && MailTo.Count > 0)
                            foreach (MailAddress em in MailTo) { mailMessage.CC.Add(em); }

                        string InfoCc = GetDisciplineInfoCc(objweb, leadlist, item, "InfoCc");
                        foreach (MailAddress em in MailTo) { mailMessage.CC.Add(em); }

                        if (!string.IsNullOrEmpty(InfoCc))
                            mailMessage.CC.Add(FormatMultipleEmailAddresses(InfoCc));

                        string IRInfoCc = GetDisciplineInfoCc(objweb, leadlist, item, DeliverableType + "InfoCc");
                        if (!string.IsNullOrEmpty(IRInfoCc))
                            MailMessageCC(IRInfoCc, mailMessage.CC);

                        string AllEmailsFromPackageList = GetPackageInfoCc(objweb, objweb.Lists["Package"], item);
                        MailMessageCC(AllEmailsFromPackageList, mailMessage.CC);

                        string InfoTo = GetDisciplineInfoCc(objweb, leadlist, item, "InfoTo");
                        if (!string.IsNullOrEmpty(InfoTo))
                        {
                            MailMessageCC(InfoTo, mailMessage.CC);
                        }

                        #region Client Emails
                        string ClientCc = "";

                        SPList Notificationlist = objweb.Lists["Notification"];
                        MailAddressCollection ClientMailTo = new MailAddressCollection();
                        MailAddressCollection ClientMailCc = new MailAddressCollection();

                        if (CheckColumnExistance(Notificationlist, "ClientCc"))
                        {
                            ClientCc = GetClientInfoCc(objweb, Notificationlist, _Notification_Name);

                            GetEmailsfromGroupSP(objweb, "", ClientCc, ref ClientMailTo, ref ClientMailCc);

                            foreach (MailAddress em in ClientMailCc) { mailMessage.CC.Add(em); }
                        }
                        #endregion

                        string HostValue = "";

                        CheckExternalExternalHost(objweb, ref HostValue);

                        subject = GetParameter(web, "ProjectName") + " - " + EmaSubject;
                        SendEmail(objweb, item, MailCC, mailMessage.CC, subject, Body, HostValue);
                    }
                    #endregion
                }
                #endregion

                LeadItem["Reviewed"] = true;
            }
            LeadItem["PartComments"] = PartComments.Text;
            LeadItem["Comment"] = CommentBox.Text;
            LeadItem["Code"] = ddlcode.SelectedValue;
            LeadItem["Status"] = "Completed";
            LeadItem["WorkflowStatus"] = "Issued to Contractor";
            LeadItem.Update();
        }        
        catch (Exception ex)
        {
            lblmsg.Text = "An error occurred. Please contact your PWS system administrator";
            SendErrorEmail(web, "Project Link: " + web.Url + "<br/>btnSubmitCloseIR_Click Error: Item Ref No: " + lblIRNo.Text + "<br/>" + ex.ToString(), ErrorSubject);
        }
        //if (web != null)
        //    web.Dispose();
    }

    protected void btnReject_Click(object sender, EventArgs e)
    {
        bool hasChecked = false;
        string RejectionTrades = "";
        string RejectionTradesMesg = "";
        foreach (GridViewRow GridRow in grdWorkflowtasks.Rows)
        {
            CheckBox CheckBoxWTask = (CheckBox)GridRow.FindControl("CheckBoxWTask");
            Label lblTrade = (Label)GridRow.FindControl("lblTrade");
            TextBox txtRejTrades = (TextBox)GridRow.FindControl("txtRejTrades");
            
            if (CheckBoxWTask.Checked)
            {
                if (string.IsNullOrEmpty(txtRejTrades.Text)) {
                    RejectionTradesMesg += "Reason of Rejection is required for " + lblTrade.Text + "<br/>";
                }

                RejectionTrades += lblTrade.Text + "|" + txtRejTrades.Text + ";";
                hasChecked = true;
            }
        }
        if (!hasChecked)
        {
            lblmsg.Text = "Select one task at least from 'select to Reject'";
            return;
        }
        else if (!string.IsNullOrEmpty(RejectionTradesMesg))
        {
            lblmsg.Text = RejectionTradesMesg;
            return;
        }
        LeadItem["RejectionTrades"] = RejectionTrades;
        LeadItem["isRejected"] = true;
        LeadItem.Update();
        Response.Write("<script>window.location = '" + SPContext.Current.Web.Url + "/Lists/LeadAction" + "';</script>");
    }

    protected void FUbtn_Click(object sender, EventArgs e)
    {
        if (IRFileUpload.HasFile)
        {
            if (Path.GetExtension(IRFileUpload.FileName).ToLower() == ".pdf")
            {
                SPWeb objweb = null;
                byte[] FinalpdfFile = IRFileUpload.FileBytes;
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                { objweb = site.OpenWeb(); }

                objweb.AllowUnsafeUpdates = true;
                lblmsg.Text = "";
         
                SPListItem LeadItem = null;

                Guid ListId = new Guid(HttpContext.Current.Request.QueryString["ListId"]);
                SPList objlistLeadTasks = objweb.Lists[ListId];

                if (isMain)
                {                                        
                    SPQuery spquery = new SPQuery();
                    spquery.Query = "<Where><Eq><FieldRef Name='Reference' /><Value Type='Text'>" + StrFullRef + "</Value></Eq></Where>";
                    SPListItemCollection colitems = objweb.Lists[LeadTaskList].GetItems(spquery);
                    if (colitems.Count == 1)
                        LeadItem = colitems[0];                    
                }

                else
                    LeadItem = objlistLeadTasks.GetItemById(int.Parse(HttpContext.Current.Request.QueryString["ID"]));      

                string attachName = lblIRNo.Text + "-Commented.pdf";

                //Delete old Attachement
                if (LeadItem.Attachments.Count > 0)
                {
                    try { LeadItem.Attachments.DeleteNow(attachName); } catch { }
                }

                LeadItem.Attachments.AddNow(attachName, FinalpdfFile);

                HyperLinkLeadAttach.NavigateUrl = LeadItem.Attachments.UrlPrefix + attachName;
                HyperLinkLeadAttach.Text = attachName;
            }
            else lblmsg.Text = "Only PDF File is allowed to upload as Final Comment";
        }
    }

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


