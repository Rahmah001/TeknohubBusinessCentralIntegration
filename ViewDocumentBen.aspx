<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ViewDocumentBen.aspx.vb" Inherits="AiicoAssistance.ViewDocumentBen" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">

        .auto-style1 {
            width: 123px;
        }
        .auto-style2 {
            width: 123px;
            height: 23px;
        }
        .auto-style3 {
            height: 23px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
   <div>
        <table >
            <tr>
                <td>
    
        <asp:Label ID="lblmsg" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
                </td>
            </tr>
            </table>

        <table style="width:100%;">
            <tr>
                <td class="auto-style1">P.I.N</td>
                <td>
                    <asp:Label ID="lblPin" runat="server"></asp:Label>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td class="auto-style2">Request Type:</td>
                <td class="auto-style3">
                    <asp:Label ID="lblUpdateType" runat="server"></asp:Label>
                </td>
                <td class="auto-style3"></td>
            </tr>
            <tr>
                <td class="auto-style1">&nbsp;</td>
                <td>&nbsp;</td>
                <td>&nbsp;</td>
            </tr>
        </table>

    </div>
    </form>
    <div id="divGallary" runat ="server" >
    </div>
</body>
</html>
