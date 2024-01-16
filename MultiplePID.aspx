<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="MultiplePID.aspx.vb" Inherits="AiicoAssistance.MultiplePID" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
    <style type="text/css">


        .auto-style1 {
            height: 26px;
        }
        </style>

        <script language="javascript" type="text/javascript">
            function validateFile(ctrlname) {
                var flag = 0;
                var uploadPath = document.getElementById(ctrlname.id).value;
                if (uploadPath == "") {
                    alert('Please browse a file to upload');
                    return false;
                }
                else {
                    if (uploadPath.indexOf('.') != -1) {
                        var validExtensions = new Array();
                        validExtensions[0] = 'pdf';
                        validExtensions[1] = 'jpeg';
                        

                        var ext = uploadPath.substring(uploadPath.lastIndexOf('.') + 1).toLowerCase();
                        for (var i = 0; i < validExtensions.length; i++) {
                            if (ext == validExtensions[i]) {
                                flag = 1;
                                break;
                            }
                        }

                        if (flag == 0) {
                            alert('The type of file is not allowed, Upload jpg,pdf,bmp or gif file');

                            return false;
                        }
                    }

                    else {
                        return false;
                    }
                }
            }
</script> 
    </head>
<body>
    <form id="form1" runat="server" enctype="multipart/form-data">
    <div>
    
        <asp:Label ID="lblmsg" runat="server" Font-Bold="True" ForeColor="Red"></asp:Label>
        <table style="width:100%;">
            <tr>
                <td class="auto-style1">
                    Temporary ID&nbsp;</td>
                <td class="auto-style1">
                    <asp:TextBox ID="txtno" runat="server" ReadOnly="True" Width="387px"></asp:TextBox>
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr id="rForm" runat="server" visible ="True" >
                <td class="auto-style1">
                    Scanned Rsa Form</td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fupForm" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cvform" runat="server" ControlToValidate="fupForm"
ErrorMessage="Please upload only pdf file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupForm)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td class="auto-style1"></td>
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
    <p>
        &nbsp;</p>
</body>
</html>
