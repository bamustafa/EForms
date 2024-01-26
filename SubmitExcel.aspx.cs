using System;
using System.Web.UI;
using System.Data;
using System.Collections;
using Microsoft.SharePoint;
using System.IO;
using Winnovative.ExcelLib;
using System.Net.Mail;
using System.Configuration;


public partial class SubmitExcel : System.Web.UI.Page
{
    #region Global Variables
    public string WinnovativeLicenseKey = "NB8GFAYHFAQUAhoEFAcFGgUGGg0NDQ0=";
    public string serverPath = @"C:\Windows\Temp\";
     //@"C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\12\TEMPLATE\LAYOUTS\C4\CDS\C4FileNameChecker.xml";
    public string Filename = default(string);


    string[] FileNameParts;

    bool flag = false;
    bool Cflag = false;
    bool ExcelListFlag = false;
    bool flagDataTable = false;

    
    SPListItem item = null;
    SPListItem RLODItem = null;
    SPFile file = null;
    public string RedirectURL = default(string);

    ArrayList CheckIfFalseExist = new ArrayList();  
    DataSet ds = new DataSet();
    DataTable dt = new DataTable("ExcelTable");

    string DateFormat = "dd/MM/yyyy";
    string GlobalError = "";
    #endregion

    protected void Page_PreInit(object sender, EventArgs e)
    {
        string MasterPageUrl = ConfigurationManager.AppSettings["MasterPageUrl"];
        if (string.IsNullOrEmpty(MasterPageUrl))
            MasterPageUrl = "~/_layouts/SAEI.master";
        else MasterPageUrl = "~/_layouts/15/SAEI.master";
        this.MasterPageFile = MasterPageUrl;
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        SPWeb web = null;

        string Message = "",
               SubmitExcelListName = "SubmitExcel";

        string HtmlTableError = "<table border='1' cellspacing='0' cellpadding='2' width='100%'><tr>" +
                                  "<td width='5%' style='text-align:center;'><b>Row</b></td>" +
                                  "<td width='25%' style='text-align:center;'><b>Filename</b></td>" +
                                  "<td width='70% style='text-align:center;'><b>Status</b></td>" +
                                "</tr>";

        try
        {
            #region VARIABLE DECLARATION
            WinnovativeLicenseKey = ConfigurationManager.AppSettings["WinnovativeLicenseKey"];
            if (string.IsNullOrEmpty(WinnovativeLicenseKey))
                WinnovativeLicenseKey = "NB8GFAYHFAQUAhoEFAcFGgUGGg0NDQ0=";

            SPList SubmitExcel;
            SPSecurity.RunWithElevatedPrivileges(delegate()
            {
                using (SPSite site = new SPSite(SPContext.Current.Web.Site.ID))
                {
                    web = site.OpenWeb();
                }
            });
            web.AllowUnsafeUpdates = true;

            SubmitExcel = web.Lists[SubmitExcelListName];

            int ID = int.Parse(Context.Request["ID"].ToString());
            try { item = SubmitExcel.GetItemById(ID); }
            catch (Exception ex) { if (ex.Message.ToLower().Contains("item does not exist")) return; }

            //item = TentativeLOD.GetItemById(ID);
            file = item.File;
            if (file == null)
            {
                MsgLabel.Text = "Kindly Select a file and try agaian";
                RedirectURL = SPContext.Current.Web.Url + "/" + SubmitExcelListName;
                return;
            }
            RedirectURL = SPContext.Current.Web.Url + "/" + file.ParentFolder.Url;
            #endregion

            #region WRITE EXCEL TEMPFILE
            Random rand = new Random();
            int OpenBookNum = rand.Next(0, 10000000);
            int dPos = file.Name.IndexOf(".");
            string excExt = file.Name.Remove(0, dPos).Trim().ToLower();

            while (File.Exists(serverPath + OpenBookNum.ToString() + excExt))
            {
                OpenBookNum = rand.Next(0, 1000000);
            }
            File.WriteAllBytes(serverPath + OpenBookNum.ToString() + excExt, file.OpenBinary());
            String dataFilePath = serverPath + OpenBookNum.ToString() + excExt;
            System.IO.FileStream sourceXlsDataStream = new System.IO.FileStream(dataFilePath, System.IO.FileMode.Open);

            ExcelWorkbook tempWorkbook = new ExcelWorkbook(sourceXlsDataStream);
            tempWorkbook.LicenseKey = WinnovativeLicenseKey;

            ExcelWorksheet sheet = tempWorkbook.Worksheets[0];
            ExcelRange range = sheet.UsedRange;
            #endregion

            #region EXCEL CELL STYLES
            ExcelCellStyle errStyle = null;
            ExcelCellStyle CStyle = null;

            for (int i = 0; i < tempWorkbook.Styles.Count; i++)
            {
                if (tempWorkbook.Styles[i].Name == "ErrorStyle")
                {
                    errStyle = tempWorkbook.Styles[i];
                    flag = true;
                }
                else if (tempWorkbook.Styles[i].Name == "CompStyle")
                {
                    CStyle = tempWorkbook.Styles[i];
                    Cflag = true;

                }
            }
            if (flag == false)
            {
                errStyle = tempWorkbook.Styles.AddStyle("ErrorStyle");
                errStyle.Font.Bold = true;
                errStyle.Alignment.VerticalAlignment = ExcelCellVerticalAlignmentType.Center;
                errStyle.Alignment.HorizontalAlignment = ExcelCellHorizontalAlignmentType.Left;
                errStyle.Fill.FillType = ExcelCellFillType.SolidFill;
                errStyle.Fill.SolidFillOptions.BackColor = System.Drawing.Color.Red;
                errStyle.Borders[ExcelCellBorderIndex.Bottom].LineStyle = ExcelCellLineStyle.Thin;
                errStyle.Borders[ExcelCellBorderIndex.Top].LineStyle = ExcelCellLineStyle.Thin;
                errStyle.Borders[ExcelCellBorderIndex.Left].LineStyle = ExcelCellLineStyle.Thin;
                errStyle.Borders[ExcelCellBorderIndex.Right].LineStyle = ExcelCellLineStyle.Thin;
                errStyle.Borders[ExcelCellBorderIndex.DiagonalDown].LineStyle = ExcelCellLineStyle.None;
                errStyle.Borders[ExcelCellBorderIndex.DiagonalUp].LineStyle = ExcelCellLineStyle.None;
            }
            if (Cflag == false)
            {
                CStyle = tempWorkbook.Styles.AddStyle("CompStyle");
                CStyle.Font.Bold = true;
                CStyle.Alignment.VerticalAlignment = ExcelCellVerticalAlignmentType.Center;
                CStyle.Alignment.HorizontalAlignment = ExcelCellHorizontalAlignmentType.Left;
                CStyle.Fill.FillType = ExcelCellFillType.SolidFill;
                CStyle.Fill.SolidFillOptions.BackColor = System.Drawing.Color.GreenYellow;
                CStyle.Borders[ExcelCellBorderIndex.Bottom].LineStyle = ExcelCellLineStyle.Thin;
                CStyle.Borders[ExcelCellBorderIndex.Top].LineStyle = ExcelCellLineStyle.Thin;
                CStyle.Borders[ExcelCellBorderIndex.Left].LineStyle = ExcelCellLineStyle.Thin;
                CStyle.Borders[ExcelCellBorderIndex.Right].LineStyle = ExcelCellLineStyle.Thin;
                CStyle.Borders[ExcelCellBorderIndex.DiagonalDown].LineStyle = ExcelCellLineStyle.None;
                CStyle.Borders[ExcelCellBorderIndex.DiagonalUp].LineStyle = ExcelCellLineStyle.None;
            }
            #endregion

            SetDataPerColumn(web, "InspType", "", "", "", range, 0);

            SetDataPerColumn(web, "Discipline", "InspType", "", "", range, 1);

            SetDataPerColumn(web, "InspPurpose", "InspType", "Discipline", "", range, 2);

            SetDataPerColumn(web, "InspDesc", "InspType", "Discipline", "InspPurpose", range, 3);
        }
        catch (Exception ex)
        {
            if (!Page.IsPostBack)
            {
                MsgLabel.Text = ex.ToString();
                SendErrorEmail(web, "E-Inspection Error<br /> SiteUrl: <br/>" + web.Url + "<br/><br/>" +
                               "ListName: <br/>" + SubmitExcelListName + "<br/><br/>" +
                               "Message: <br/>" + ex.Message + "<br/><br/>" +
                               "StackTrace: <br/>" + ex.StackTrace);
            }
        }
        finally { if (web != null) web.Dispose(); }
    }
    
    protected void LinkButton1_Click(object sender, EventArgs e)
    {
       Response.Redirect(RedirectURL);
    }

    public static void SendErrorEmail(SPWeb web, string Body)
    {
        string from = GetParameter(web,"FromError");

        string to = "";
        string ErrorAdmin = GetParameter(web,"ErrorAdmin");
        if (string.IsNullOrEmpty(ErrorAdmin))
            to = GetParameter(web,"AdminEmails");
        else to = ErrorAdmin;

        string subject = Environment.MachineName + " - " + GetParameter(web,"ProjectName") + " - Error on SubmitLOD.aspx.cs ";
        string body = Body;

        SmtpClient myClient = new SmtpClient(ConfigurationManager.AppSettings["SmtpServer"]);

        if (!string.IsNullOrEmpty(GetParameter(web,"FromErrorSmtpUser")))
            myClient.Credentials = new System.Net.NetworkCredential(GetParameter(web,"FromErrorSmtpUser"), GetParameter(web,"FromErrorSmtpPWD"));
        else
            myClient.UseDefaultCredentials = true;

        MailMessage msg = new MailMessage();
        msg.From = new MailAddress(from, GetParameter(web,"FromErrorName"));
        msg.To.Add(to);
        msg.Subject = subject;
        msg.Body = body;
        msg.IsBodyHtml = true;
        myClient.Send(msg);
    }

    public static string GetParameter(SPWeb web, string key)
    {
        string retValue = "";
        try
        {
            SPList list = web.Lists["Parameters"];
            SPQuery spquery = new SPQuery();
            spquery.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" +
                            key + "</Value></Eq></Where>";
            SPListItemCollection colitems = list.GetItems(spquery);
            if (colitems.Count == 1)
            {
                SPListItem item = colitems[0];
                if(item["Value"] != null && item["Value"].ToString() != "")
                 retValue = item["Value"].ToString().Trim();
            }
            else if (colitems.Count == 0)
            {
                if (key == "LOD-Columns-InternalName" || key == "LOD-Columns-DisplayName" || key == "LOD-Columns-DataType" || key == "LOD-Required-Columns" ||
                    key == "RevStartsWith" || key == "ContractorCode" || key == "ContractorCodeColumn" || key == "ErrorAdmin" || key =="Lookup" || key == "Case_Sensative" || key == "UpdateTitle")
                    return "";
                else
                    SendErrorEmail(web, "SiteUrl: <br/>" + web.Url + " <br/><br/>" + "Error Message:<br/> Error in GetParameter: No records for Type " + key);
            }
            else
            {
                SendErrorEmail(web, "SiteUrl: <br/>" + web.Url + " <br/><br/>" +
                               "Error Message:<br/> Error in GetParameter: Multiple records for Type " + key);
            }
            return retValue;

        }
        catch (Exception e)
        {
            SendErrorEmail(web, "SiteUrl: <br/>" + web.Url + "<br/><br/>" +
                           "Error Message:<br/> Error in GetParameter Type:" + key + " <br/><br/> " +
                           "Message: <br/>" + e.Message + "<br/><br/>" +
                           "StackTrace: <br/>" + e.StackTrace);
            return "";
        }
    }

    public void SetStatusMessage(ExcelRange range, int index, string Message, ExcelCellStyle Style, bool ExcelListFlag, ref ArrayList CheckIfFalseRecord, string Filename)
    {
        range.Columns[3].Cells[index].Style = Style;
        range.Columns[3].Cells[index].Text = Message;
        CheckIfFalseRecord.Add(ExcelListFlag);
        CheckIfFalseExist.Add(ExcelListFlag);

        int row = index + 1;
        GlobalError += "<tr>" +
                         "<td style='vertical-align:middle;text-align:center;'>" + row + "</td>" +
                         "<td style='vertical-align:middle;'>" + Filename + "</td>" +
                         "<td style='vertical-align:middle;color:red'>" + Message + "</td>" +
                       "</tr>";
    }

    public static void SetDataPerColumn(SPWeb web, string listName, string Column_Name1, string Column_Name2, string Column_Name3, ExcelRange range, int index)
    {
        string ColumnVal = "";
        bool isVisited = false;

        for (int i = 1; i < range.Rows.Count; i++)
        {
            ColumnVal = range.Columns[index].Cells[i].Text;

            if (!isVisited)
            {
                CreateList(web, listName, Column_Name1, Column_Name2, Column_Name3);
                isVisited = true;
            }

            SPList newList = web.Lists[listName];
            SPQuery Query = new SPQuery();
            Query.Query = "<Where>" +
                             "<Eq><FieldRef Name='Title' /><Value Type='Text'>" + ColumnVal + "</Value></Eq>" +
                         "</Where>";
            SPListItemCollection items = newList.GetItems(Query);
            SPListItem item = null;

            if (items.Count == 0)
                item = items.Add();
            else item = items[0];
             
            item["Title"] = ColumnVal;

            if (index == 1)
              SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name1, 0, ref item); //LEVEL ONE (INSPECTION TYPE)

            if (index == 2)
            {
                //LEVEL ONE (INSPECTION TYPE)
                SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name1, 0, ref item);

                //LEVEL TWO (DISCIPLINE)
                SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name2, 1, ref item);
            }

            if (index == 3)
            {
                //LEVEL ONE (INSPECTION TYPE)
                SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name1, 0, ref item);

                //LEVEL TWO (DISCIPLINE)
                SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name2, 1, ref item);

                //LEVEL TWO (INSP_PURPOSE)
                SetMultiLookup_Values(web, range, index, ColumnVal, Column_Name3, 2, ref item);
            }

            if (item != null)
                item.Update();
        }
    }

    public static void CreateList(SPWeb web, string Listname, string Column_name1, string Column_name2, string Column_name3)
    {
        SPList isListExist = web.Lists.TryGetList(Listname);
        if (isListExist == null)
        {
            SPListCollection lists = web.Lists;
            lists.Add(Listname, "", SPListTemplateType.GenericList);
        }

        SPList newList = web.Lists[Listname];
        string AcronymField = "Acronym";
        if (!newList.Fields.ContainsField(AcronymField))
        {
            newList.Fields.Add(AcronymField, SPFieldType.Text, false);
            SPView view = newList.DefaultView;
            view.ViewFields.Add(AcronymField);
            view.Update();
        }

        if (!string.IsNullOrEmpty(Column_name1))
        {
            if (!newList.Fields.ContainsField(Column_name1))
                CreateSPLookup(web, newList, Column_name1);
        }

        if (!string.IsNullOrEmpty(Column_name2))
        {
            if (!newList.Fields.ContainsField(Column_name2))
                CreateSPLookup(web, newList, Column_name2);
        }

        if (!string.IsNullOrEmpty(Column_name3))
        {
            if (!newList.Fields.ContainsField(Column_name3))
                CreateSPLookup(web, newList, Column_name3);
        }
}

    public static void CreateSPLookup(SPWeb web, SPList newList, string Column_Name)
    {
        SPList PrevListname = web.Lists[Column_Name];

        newList.Fields.AddLookup(Column_Name, PrevListname.ID, false);
        SPFieldLookup lkp = (SPFieldLookup)newList.Fields[Column_Name];
        lkp.AllowMultipleValues = true;
        lkp.LookupField = newList.Fields["Title"].InternalName;
        lkp.Update();

        // make new column visible in default view
        SPView view = newList.DefaultView;
        view.ViewFields.Add(Column_Name);
        view.Update();
    }

    public static void GetLookFieldIDS(string lookupValue, SPList lookupSourceList, ref int id)
    {
        //SPFieldLookupValueCollection lookupIds = new SPFieldLookupValueCollection();
        SPQuery query = new Microsoft.SharePoint.SPQuery();
        query.Query = "<Where><Eq><FieldRef Name='Title' /><Value Type='Text'>" + lookupValue + "</Value></Eq></Where>";
        SPListItemCollection listItems = lookupSourceList.GetItems(query);
        foreach (SPListItem item in listItems)
        {
            id = item.ID;
        }
    }

    public static void SetMultiLookup_Values(SPWeb web, ExcelRange range, int index, string CurrentValue, string Column_Name, int TargetIndex, ref SPListItem item)
    {
        SPFieldLookupValueCollection values = new SPFieldLookupValueCollection();
        for (int j = 1; j < range.Rows.Count; j++)
        {
            string NewVal = range.Columns[index].Cells[j].Text;
            if (NewVal == CurrentValue)
            {
                int id = 0;
                string val = range.Columns[TargetIndex].Cells[j].Text;

                GetLookFieldIDS(val, web.Lists[Column_Name], ref id);
                if (id != 0)
                {
                    SPFieldLookupValue lkpval = new SPFieldLookupValue(id, val);
                    values.Add(lkpval);
                }
            }
        }
        if (values != null && values.Count > 0)
            item[Column_Name] = values;
    }

}



