<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="NewUploadDocument.aspx.vb" Inherits="AiicoAssistance.NewUploadDocument" %>

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
                        validExtensions[2] = 'bmp';
                        validExtensions[3] = 'gif';
                        validExtensions[4] = 'jpg';


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
                <td class="auto-style1">PIN:</td>
                <td class="auto-style1">
                    <asp:TextBox ID="txtno" runat="server" ReadOnly="True" Width="387px"></asp:TextBox>
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr id="rForm" runat="server" visible ="True" >
                <td class="auto-style1">
                    <asp:Label ID="lblform" runat="server" Text="Scanned Form"></asp:Label>
                </td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fupForm" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cvform" runat="server" ControlToValidate="fupForm"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupForm)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr id="r0" runat="server" visible ="false" >
                <td class="auto-style1">
                    <asp:Label ID="Label0" runat="server" Text="Label"></asp:Label>
                </td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fupdoc0" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv0" runat="server" ControlToValidate="fupdoc1"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupdoc1)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td class="auto-style1"></td>
            </tr>
               <tr id="r1" runat="server" visible ="false">
                <td class="auto-style1">
                    <asp:Label ID="Label1" runat="server" Text="Label"></asp:Label>
                </td>
                <td class="auto-style1">
                    <asp:FileUpload ID="fupDoc1" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv1" runat="server" ControlToValidate="fupDoc2"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc2)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td class="auto-style1"></td>
            </tr>
            <tr id="r2" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label2" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc2" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv2" runat="server" ControlToValidate="fupDoc3"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc3)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r3" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label3" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc3" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv3" runat="server" ControlToValidate="fupDoc4"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc4)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r4" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label4" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc4" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv4" runat="server" ControlToValidate="fupDoc4"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc4)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r5" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label5" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc5" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv5" runat="server" ControlToValidate="fupDoc5"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc5)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r6" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label6" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc6" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv6" runat="server" ControlToValidate="fupDoc6"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc6)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r7" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label7" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc7" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv7" runat="server" ControlToValidate="fupDoc7"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc7)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r8" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label8" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc8" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv8" runat="server" ControlToValidate="fupDoc8"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc8)" ForeColor="Red"></asp:CustomValidator>
                </td>
                <td>&nbsp;</td>
            </tr>
            <tr id="r9" runat="server" visible ="false">
                <td>
                    <asp:Label ID="Label9" runat="server" Text="Label"></asp:Label>
                </td>
                <td>
                    <asp:FileUpload ID="fupDoc9" runat="server" Width="532px" />
                    <asp:CustomValidator ID="cv9" runat="server" ControlToValidate="fupDoc9"
ErrorMessage="Please upload only pdf or docx file" SetFocusOnError="True" ValidateEmptyText="True" ClientValidationFunction="validateFile(fupDoc9)" ForeColor="Red"></asp:CustomValidator>
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
