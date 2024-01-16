<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="LoadImages.aspx.vb" Inherits="AiicoAssistance.LoadImages" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">
        .auto-style1 {
            height: 26px;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
    
        <asp:Label ID="lblmsg" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
        <table style="width:100%;">
            <tr>
                <td class="auto-style1">No:</td>
                <td class="auto-style1">
                    <asp:TextBox ID="txtno" runat="server" ReadOnly="True" Width="387px"></asp:TextBox>
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr>
                <td class="auto-style1">Passport Path:</td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fuppassport" runat="server" Width="532px" />
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr>
                <td class="auto-style1">Signature Path:</td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fupSignature" runat="server" Width="532px" />
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr>
                <td>Right Thumb Path:</td>
                <td>
                    <asp:FileUpload ID="fupRightthumb" runat="server" Width="532px" />
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>Left Thumb Path:</td>
                <td>
                    <asp:FileUpload ID="fupleftthumb" runat="server" Width="532px" />
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td>&nbsp;</td>
                <td>
                    <asp:Button ID="btnsave" runat="server" Text="Save &amp; Return" Width="157px" />
                </td>
                <td>&nbsp;</td>
            </tr>
        </table>
    
    </div>
    </form>
</body>
</html>
