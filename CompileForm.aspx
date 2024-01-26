<%@ Page Language="C#" AutoEventWireup="true" CodeFile="CompileForm.aspx.cs" Inherits="CompileForm" EnableSessionState="True" Async="true"%>

<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %> 
<%@ Import Namespace="Microsoft.SharePoint.WebControls" %> 
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Register TagPrefix="SP" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=14.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>

<!DOCTYPE html>

<html> 

<head runat="server"> <%--runat="server"--%>

<meta charset="utf-8" />
<meta name="viewport" content="width=device-width, initial-scale=1" http-equiv="X-UA-Compatible">

<style type="text/css"> </style>
<script type="text/javascript" src="controls/preloader/jquery-1.9.0-vsdoc.js"></script>
<script type="text/javascript" src="controls/preloader/jquery-1.9.0.js"></script>
<script type="text/javascript" src="controls/preloader/jquery.dim-background.min.js"></script>

<link rel="stylesheet" type="text/css"  href="MobileCss/w3.css" />
<link rel="stylesheet" type="text/css"  href="MobileCss/w3-theme-dark-grey.css" />
<link rel="stylesheet" type="text/css"  href="MobileCss/font-awesome/css/font-awesome.min.css" />

<style>
    .center {
          position: absolute;
          top: 400px;
          left: 900px;
          width:100px;
          height:100px;
          visibility:hidden;
        }
    .showcase{
        background:url('images/Inspc.jpg');
        background-position:center;
        height:250px;     
        padding:90px 10px;
        color:#fff;      
        background-repeat: no-repeat;
        background-size: cover;
    }
    .showcase h1{
        font-size:42px;
        font-weight:800;
        text-shadow: 4px 4px 6px #000000;        
    }
    .fa{
        font-size:17px;
    }
    .showcase hr{
        width:27%;
        margin:auto;
        border-width:3px;
        border-color:#f44336;
        margin-top:30px; 
        /*padding-left: 380px;*/       
    }
    .fillBody hr{
        width:80%;
        margin:auto;
        border-width:2px;
        border-color:Gray;
    }
    .fillBody{
        background-color: #fafafa; /*floralwhite*/ /*aliceblue*/
        background-position:center;
        /*height:900px;*/
        background-size: cover;         
        padding:30px 10px;       
        color:#fff;
    }
    .ColumnCSS
        {
            height:25px;
            Width: 10%;
        	font-family: Verdana;
        	font-size:13px;         
            background-color: #3CDBC0; /*#73737e*/
            color: black; /*white*/
            font-weight: bold;	
            vertical-align: middle;
            /*border: 1px solid #5d5d6b !important;*/
            /*border-collapse: collapse;*/
        }
    .ValueCSS 
       {
        height:25px;
       	font-family: Verdana;
       	font-size: 13px;
       	background-color: white;
       	/*font-weight: bold;*/
        /*border: 1px solid #5d5d6b !important;
        border-collapse: collapse;*/
       }
    .textcss {
            color:black;
            /*font-weight:bold;*/
            font-family:Verdana, Arial, sans-serif;
            font-size:12px;
            height:35px;
        }
    .auto-style1 {
        width: 100%;
        border-collapse: collapse;
    }
    .auto-style2 {
        width: 100%;
        /*border-collapse: collapse;*/
    }
    .ReadBox {              
        box-shadow: 4px 4px 2px #aaaaaa;
        border: 1px solid;
        border-color: #aaaaaa;
        color: black;
        padding: 5px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        margin: 4px 2px;
        cursor: not-allowed;
        border-radius: 8px;      
}
    .ReadBoxwithnoCursor {              
        box-shadow: 4px 4px 2px #aaaaaa;
        border: 1px solid;
        border-color: #aaaaaa;
        color: black;
        padding: 5px 20px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        margin: 4px 2px;
        border-radius: 8px;      
}
    .Paragraphs {
      background-color: #fafafa;
      box-shadow: 4px 4px 3px #aaaaaa;
      color: black;
      font-family: Verdana;
      font-size: 13px;
      font-weight:bold;
      padding: 4px 18px;
      text-align: center;
      text-decoration: none;
      display: inline-block;
      margin: 4px 2px;  
      border-radius: 8px;
      cursor: not-allowed;
      font-size:13px;  
      border: 1px solid;
      border-color: #aaaaaa;         
}
    .EntryParg{
        background-color: #fff;
        box-shadow: 4px 4px 2px #aaaaaa;
        border: 1px solid;
        border-color: #aaaaaa;
        color: black;
        padding: 5px 20px;  
        text-decoration: none;
        display: inline-block;
        margin: 4px 2px;
        border-radius: 8px;
    }
    .button {
        background-color: #f1f1f1; /*teal*/
        border: 1px solid;
        border-color: #aaaaaa;
        color: #37c1aa; /*#fff*/
        font-family: Verdana;
        font-size: 13px;
        padding: 6px 18px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 8px;
        box-shadow: 4px 4px 3px #aaaaaa;
        font-weight:bold;
}

.button:hover {
  background-color: #3CDBC0; /*#f1f1f1*/
  color: #fff; /*black*/
}
.buttonCancel {
        background-color: #f1f1f1; /*red*/
        border: 1px solid;
        border-color: #aaaaaa;
        color: red; /*#fff*/
        font-family: Verdana;
        font-size: 13px;
        padding: 6px 18px;
        text-align: center;
        text-decoration: none;
        display: inline-block;
        margin: 4px 2px;
        cursor: pointer;
        border-radius: 8px;
        box-shadow: 4px 4px 3px #aaaaaa;
        font-weight:bold;
}

.buttonCancel:hover {
  background-color: red; /*#f1f1f1*/
  color: #fff; /*red*/
}
.ChkBoxClass input {width:15px; height:15px; margin-right: 5px; vertical-align: middle;}

 .rounded-corners
{    
    border:0.5px solid !important; 
    border-radius: 15px;   
    overflow: hidden;
} 

 .rounded-corners-First
{    
    border:1px solid #000000 !important; 
    border-radius: 15px;
    border-width: 2px 4px;   
} 

 .file-panel {
    display: none;
    position: absolute;
    top: 0;
    left: 0;
    width: 300px;
    height: 200px;
    background-color: white;
    border: 1px solid #ccc;
    padding: 10px;
    box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    z-index: 1000;
}
</style>

    <script>    
    $(document).ready(function () {
        $('input[type=submit]').click(function () {
            $(this).dimBackground({
                darkness: 0.1
            }, function () {
                $('#loader').css("visibility", "visible").dimBackground({ darkness: 0.8 });
            });
        });
    });
    </script>

</head>

<body>

<form id="form1" runat="server">
    <asp:Image runat="server" ID="loader" ImageUrl="~/Images/loading1.gif" CssClass="center" />
    <section class="showcase">
        <div class="w3-container w3-center">
            <h1 class="w3-text-shadow w3-animate-opacity"><%= Category %> Compilation Form</h1> <%--E-Inspection Review by Lead Action--%>
            <hr />
        </div>
    </section> 

    <section class="fillBody">
        <div class="w3-container w3-center"> 
            
            <div class="ReadBoxwithnoCursor w3-animate-opacity" style="width:100%;">   
        
            <table class="rounded-corners-First auto-style2" style="margin-top:20px;" align="center" cellpadding="0px" cellspacing="0px">                
                <tr>
                    <td class="ColumnCSS" style="border-top-left-radius: 15px;border-bottom:1px solid !important;border-right:1px solid !important;">Reference</td>
                    <td class="ColumnCSS" style="border-top-right-radius: 15px;border-bottom:1px solid !important;">Discipline:</td>                   
                </tr>
                <tr>
                    <td class="ValueCSS" style="border-bottom:1px solid !important;border-right:1px solid !important;">
                        <asp:Label ID="lblIRNo" runat="server" class="textcss" ></asp:Label>
                    </td>
                    <td class="ValueCSS" style="border-bottom:1px solid !important;"> 
                        <asp:Label ID="lblTrade" class="textcss" runat="server" ></asp:Label>
                    </td>                 
                </tr>
                <tr>                    
                    <td class="ColumnCSS" style="border-bottom:1px solid !important;border-right:1px solid !important;">CM Representative</td>
                    <td class="ColumnCSS" style="border-bottom:1px solid !important;">E-Inspection Attachment&nbsp;:</td>
                </tr>
                <tr>                    
                    <td class="ValueCSS" style="border-bottom:1px solid !important;border-right:1px solid !important;"><asp:Label ID="lblRE" class="textcss" runat="server" ></asp:Label></td>
                    <td class="ValueCSS" style="border-bottom:1px solid !important;"><asp:HyperLink ID="Pdfhyplink" style="color:#6083e0;" runat="server" Target="_blank" ></asp:HyperLink></td>
                </tr>
                <tr>
                    <td class="ColumnCSS" style="border-bottom:1px solid !important;" colspan="2">Title</td>
                </tr>
                <tr>
                    <td class="ValueCSS" style="border-bottom-left-radius: 15px;border-bottom-right-radius: 15px;" colspan="2"><asp:Label ID="lblRIWTitle" class="textcss" runat="server" ></asp:Label></td>
                </tr>               
            </table>           
             
            <div class = "rounded-corners" style="margin-top:20px;"> 
            <asp:GridView ID="grdWorkflowtasks" width="100%" runat="server" AutoGenerateColumns="False" OnRowDataBound="grdWorkflowtasks_RowDataBound" CellPadding="1" BorderWidth="0.5px" EnableModelValidation="True" ForeColor="Black" GridLines="Vertical" BackColor="White">               
                <HeaderStyle Height="26px"/>
                <AlternatingRowStyle BackColor="#f2f2f2" />
                <FooterStyle BackColor="White" ForeColor="#333333" />
                <RowStyle Height="32px" BackColor="White" ForeColor="#333333" />
                    <Columns >       
                        <asp:TemplateField  HeaderText="Select to Reject" Visible="true" ItemStyle-Width="5%">
                            <ItemTemplate>
                                <asp:CheckBox ID="CheckBoxWTask" CssClass="ChkBoxClass" runat="server" Visible="true" Checked="false" ></asp:CheckBox>                           
                            </ItemTemplate>
                       </asp:TemplateField>
                          
                        <asp:TemplateField HeaderText="ID"  Visible="False"  >
                            <ItemTemplate>
                                <asp:Label ID="lblID" runat="server" Text='<%#Eval("ID")%>'></asp:Label>
                            </ItemTemplate>
                       </asp:TemplateField>

                    <asp:TemplateField HeaderText="Title" ItemStyle-Width="10%"  Visible="False" >
                        <ItemTemplate >
                            <asp:Label ID="lblTitle" runat="server" Text='<%#Eval("Title")%>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Trade"  ItemStyle-Width="6%" >
                        <ItemTemplate>
                            <asp:Label ID="lblTrade" runat="server" Text='<%# Eval("Trade") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>


                    <asp:TemplateField HeaderText="AssignedTo"  ItemStyle-Width="10%" >
                        <ItemTemplate>
                            <asp:Label ID="lblAssignedTo" runat="server" Text='<%# Eval("AssignedTo") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Date" ItemStyle-Width="5%">
                        <ItemTemplate>
                            <asp:Label ID="lblStartDate" runat="server" Text='<%# Eval("Date","{0:d}") %>'></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>                   

                    <asp:TemplateField HeaderText="Status" ItemStyle-Width="5%">
                        <ItemTemplate>
                            <asp:Label ID="lblStatus" runat="server" Text='<%# Eval("Status") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Code" ItemStyle-Width="8%">
                        <ItemTemplate>
                            <asp:Label ID="lblPartCode" runat="server" Text='<%# Eval("Code") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Status" Visible="false">
                        <ItemTemplate>
                            <asp:Label ID="lblWTStatus" runat="server" Text='<%# Eval("Status") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>


                    <asp:TemplateField HeaderText="Comments" ItemStyle-Width="20%" >
                        <ItemTemplate>
                            <asp:Label ID="lblComments" runat="server" Text='<%# Eval("Comment") %>' ></asp:Label>
                        </ItemTemplate>
                    </asp:TemplateField>

                    <asp:TemplateField  Visible="true" ItemStyle-Width="7%" HeaderText="Include Part Comment ?">
                        <ItemTemplate>
                            <asp:CheckBox ID="IncludePartCommentCheckBox" CssClass="ChkBoxClass" runat="server" Visible="true" Checked="true" ></asp:CheckBox>                           
                        </ItemTemplate>
                    </asp:TemplateField>

                   <asp:TemplateField HeaderText="Reason of Rejection" ItemStyle-Width="20%" >
                        <ItemTemplate>
                            <asp:TextBox ID="txtRejTrades" runat="server" Text='<%# Eval("RejectionTrades") %>' width="100%" TextMode="MultiLine" Rows="5"></asp:TextBox>
                        </ItemTemplate>
                    </asp:TemplateField>


                    <asp:TemplateField HeaderText="Attachments (select to include)" ItemStyle-Width="13%">
                        <ItemTemplate>                      
                            <asp:CheckBoxList id="Chkattachments" CssClass="ChkBoxClass" runat="server" style="text-align: left;margin-left:5px" ></asp:CheckBoxList>
                        </ItemTemplate>                       
                    </asp:TemplateField>

                    <asp:TemplateField HeaderText="Preview Link" ItemStyle-Width="10%" Visible="false">
                        <ItemTemplate>
                            <asp:Label runat="server" ID="lblattachments"></asp:Label>                      
                        </ItemTemplate>
                    </asp:TemplateField>

                </Columns>
                <RowStyle ForeColor="Black" Font-Size="12px" BorderColor="black" BorderStyle="Solid" BorderWidth="1px"/>  
                <PagerStyle BackColor="#6C6C6C" ForeColor="White" HorizontalAlign="Center" Font-Bold="True"/>
                <SelectedRowStyle BackColor="#339966" Font-Bold="True" ForeColor="White" />
                <FooterStyle BackColor="White" ForeColor="#000066" />
                <HeaderStyle BackColor="#23d1b3" Font-Bold="True" ForeColor="black" Font-Size="13px" BorderColor="black" BorderStyle="Solid" BorderWidth="0.5px"/>              
            </asp:GridView>  
            </div>                       

<%--            <div class="w3-animate-opacity" style="text-align:left;width:100%;">   
                <asp:Label style="margin-top:15px;" ID="ItemsLbl" runat="server" CssClass="Paragraphs">
                    <i  class="fa fa-cog fa-spin" aria-hidden="true"></i> 
                    <span style="color: #37c1aa;">Hint:</span>
                    <span style="color:red;"> Please uncheck the Open part tasks that you dont want to include to be able to close the <%= Category %>.</span>
                </asp:Label>               
            </div>--%>
            <br />
            </div>           
            
            <div class="ReadBoxwithnoCursor w3-animate-opacity" style="text-align:left;width:100%;">
                    
            <asp:Panel runat="server" ID="PanelRE" Visible="true">               
                           
                <div class="w3-animate-opacity">   
	                <asp:Label ID="Label6" style="margin-top:20px;" visible="false" runat="server" CssClass="Paragraphs"><span style="text-decoration: underline;color: #37c1aa;">Part Comments:</span> as seen by contractor on Form in the Inspectors Report box, please modify appropriately</asp:Label>            
	                <asp:TextBox runat="server" ID="PartComments" style="margin-top:10px;" CssClass="EntryParg" visible="false" Rows="10" TextMode="MultiLine" Width="60%"> </asp:TextBox>
                </div>            

                <div class="w3-animate-opacity">   
	                <asp:Label ID="lblcode" style="margin-top:20px;text-decoration: underline;" ForeColor="#37c1aa" visible="true" runat="server" CssClass="Paragraphs">Set Final Status:</asp:Label>  
                    <asp:dropdownlist ID="ddlcode" runat="server" style="font-family:Verdana;font-size:12px;box-shadow: 2px 2px 1px #aaaaaa;padding: 4px 18px;border-radius: 8px;color:#505256">
                    <asp:ListItem Text="Select Code"></asp:ListItem>
                    <asp:ListItem Text="Approved"></asp:ListItem>
                    <asp:ListItem Text="Approved as noted"></asp:ListItem>
                    <asp:ListItem Text="Witnessed (for APM)"></asp:ListItem>
                    <asp:ListItem Text="Rejected"></asp:ListItem>
                    <asp:ListItem Text="Work is Not Ready"></asp:ListItem>
                    <asp:ListItem Text="Cancelled"></asp:ListItem>
                </asp:dropdownlist>          
                </div>                                      

               <%-- <div style="display:none">                                                              
                    <p id="upload-area" > <asp:Label ID="lblFIle" runat="server" class="textcss" Text="IR File :" ForeColor="#37c1aa"  ></asp:Label>  
                    <asp:FileUpload ID="IRFileUpload" runat="server" Width="150px"  />  
                </div>    --%>         

                <div class="w3-animate-opacity">   
	                <asp:Label ID="Label3" style="margin-top:30px;" runat="server" CssClass="Paragraphs"><span style="text-decoration:underline;color:#37c1aa;">CM's respresentative Comments:</span> as seen by contractor on the cover page, please set appropriately</asp:Label>            
	                <asp:TextBox runat="server" ID="CommentBox" style="margin-top:10px;color:#505256;" CssClass="EntryParg" Rows="10" TextMode="MultiLine" Width="100%"> </asp:TextBox>
                </div>

                <asp:RadioButtonList Visible="false" runat="server" ID="rblist_PDFpagesToInclude"  CellPadding="10" RepeatColumns="2">
                    <asp:ListItem Text ="Include all pages from contractor RIW pdf attachement" Value="AllPagesFromRIW" Selected="True" /> 
                    <asp:ListItem Text ="Only CRS form (Use when the part attachment already include the required RIW pages)" Value="CoverPageFromRIW" />
                </asp:RadioButtonList>

                <div style="display:none">                                                              
                    <asp:Label ID="Label5" runat="server" class="textcss" Text="Code   :" ForeColor="#37c1aa" >Put the Code and Comments on:</asp:Label>
                    <asp:RadioButtonList runat="server" ID="rblist_CommentsLoc"  CellPadding="10" RepeatColumns="3">
                        <asp:ListItem Text ="CRS Form" Value="CRSForm" Selected="True" /> 
                        <asp:ListItem Text ="Cover Page From RIW" Value="CoverPageFromRIW" /> 
                        <asp:ListItem Text ="Blank Page" Value="BlankPage" />
                        <asp:ListItem Text ="First Page of merged attachments" Value="PDFAttachment" />
                        <%-- <asp:ListItem Text ="Item4" Value="4" />--%>
                    </asp:RadioButtonList><br />
                    <asp:CheckBox Checked="true" runat="server" ID="ckb_ApplyWaterMark" Text="Apply Image-Signature (paraf) with date and page count on every page of compiled PDF" /><br />
                    <asp:CheckBox Checked="true" runat="server" ID="ckb_ApplyDigiSign" Text="Apply Digital Signature on compiled PDF" /><br />
                    <br /><br />
                </div>                      

            </asp:Panel>                     

            <div class="w3-animate-opacity">   
	            <asp:Label ID="Label4" style="margin-top:30px;text-decoration: underline;" runat="server" ForeColor="#37c1aa" CssClass="Paragraphs">Compiled Attachment:</asp:Label>                           
	            <strong><asp:HyperLink ID="HyperLinkLeadAttach" runat="server" style="font-family:Verdana;font-size:13px;padding: 4px 18px;border-radius: 8px;color:#6083e0" Target="_blank"></asp:HyperLink></strong>
                 <%--<asp:ImageButton runat="server" ID="imgbtnbackground" ImageUrl="Images/preview.png" OnClick="ImageButton_Click"/>--%>
                <br />
                <br />  
                <iframe ID="embedpdfViewer" visible="false" runat="server" style="width:70%; height:800px;border:1px solid #aaaaaa;box-shadow: 4px 4px 2px #aaaaaa;"></iframe>               
            </div>                        
            
            <table class="auto-style3">   
                
                <tr>
                    <td>             
                        <asp:Panel ID="FUPanel" runat="server" Visible ="false">
                            <asp:Label ID="Label7" style="margin-top:20px;text-decoration: underline;" ForeColor="#37c1aa" visible="true" runat="server" CssClass="Paragraphs">Final Reply File:</asp:Label> 
                            <asp:FileUpload ID="IRFileUpload" runat="server" Width="350px"  CssClass="button" />
                            <asp:Button ID="FUbtn" runat="server" Text="Upload" OnClick="FUbtn_Click" visible="true" CssClass="button" />
                        </asp:Panel>
                    </td>
                </tr>

            <tr align="center">                
                <td><br /><asp:Label ID="lblmsg" runat="server" style="color:red;font-family:Verdana;font-weight:bold;font-size:14px;padding: 4px 18px;border-radius: 8px;"></asp:Label>
                </td>
             </tr>
             <tr>
                <td align="center" >
                    <br />
                    <hr />

                    <asp:Button ID="btnSave" runat="server" Text="Save Code and Comments" OnClick="btnSave_Click" style="margin-top:30px;" CssClass="button"/>
                    <asp:Button ID="btnReject" runat="server" Text="Reject" OnClick="btnReject_Click" style="margin-top:30px;" CssClass="button"/>
                    <asp:Button ID="btnUpdateAREMergePdf" runat="server" Text="Compile PDF and Merge attachments" Visible="false" OnClick="btnUpdateAREMergePdf_Click" CssClass="button" />
                    <asp:Button ID="SubmitCloseIR"  runat="server" OnClick="btnSubmitCloseIR_Click" Text="Submit Final and Close" Visible="false" CssClass="button" />
                    <asp:Button ID="btnCancel" runat="server" Text="Cancel and Go Back" OnClick="btnCancel_Click" CssClass="buttonCancel"/>                    <br />                   
                    <br />
                    <asp:Button ID="btnAssignPart" Visible="false" runat="server" Text="Go to Assign Additional Part Page" OnClick="btnAssignPart_Click" style="height: 26px"/>
                    
                    </td>
            </tr>
        </table>       

            <asp:Label ID="Label2" runat="server" ForeColor="Red"></asp:Label> 
            <asp:HiddenField ID="pdflink" runat="server"   />

        </div>       
        </div>               
    </section>       
  </form>
</body>
</html>

