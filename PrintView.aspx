<%@ Page Language="C#" AutoEventWireup="true" CodeFile="PrintView.aspx.cs" Inherits="PrintView" EnableSessionState="True" Debug="true" %>

<asp:Content ID="Content4" ContentPlaceHolderId="PlaceHolderTitleBreadcrumb" runat="server">
    <asp:SiteMapPath SiteMapProvider="SPContentMapProvider" id="ContentMap" SkipLinkText="" runat="server"/>
</asp:Content>

<asp:Content ID="Content1" ContentPlaceHolderId="PlaceHolderMain" runat="server">
    <script type="text/javascript">
        function PrintCustom() {
            var patt = /.+_FormControl.+__ViewContainer/gi;
            var alldivs = document.getElementsByTagName('div');
            var printpageHTML = '';
            for (var i = 0; i < alldivs.length; i++) {
                if (patt.test(alldivs[i].id)) {
                    printpageHTML = '<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN">\r\n <HTML><HEAD>\n' +
                                    document.getElementsByTagName('HEAD')[0].innerHTML +
                                  '</HEAD>\n<BODY>\n' +
                                  alldivs[i].innerHTML.replace('inline-block', 'block') +
                                  '\n</BODY></HTML>';
                    break;
                }
            }
            var printWindow = window.open('', 'printWindow');
            printWindow.document.open();
            printWindow.document.write(printpageHTML);
            printWindow.document.close();
            printWindow.print();
        }
    </script>


    <input onclick="javascript: PrintCustom();" type="button" value="Print"/>
    <asp:Label runat="server" ID="Formlbl" Text="" ></asp:Label>
</asp:Content>