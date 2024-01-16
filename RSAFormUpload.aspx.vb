Imports System.IO
Imports System.Data.SqlClient

Public Class RSAFormUpload
    Inherits System.Web.UI.Page


    Public pathname As String
    Public docnames As String
    Dim docObj As New ArrayList
    Dim OptionalArray As New ArrayList
    Dim Nos As String
    Dim PIN As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        writefailure("")
        If Request.QueryString("Path") <> Nothing Then
            pathname = Request.QueryString("Path")
        End If


        If Request.QueryString("PIN") <> Nothing Then
            PIN = Request.QueryString("PIN")
        End If

        If Request.QueryString("nos") <> Nothing Then
            Nos = Request.QueryString("nos")
        End If
        
        Dim query As String

        query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor Rsa Form Path] with (nolock) where [Code] ='" & Nos & "'"
        Dim ds As DataSet = returndataset(query)
        If ds.Tables(0).Rows.Count > 0 Then
            rForm.Visible = False
            writefailure("RSA form previously attached for this record.")

            btnsave.Visible = False
            Exit Sub
        End If
        If PIN Is Nothing Then PIN = ""

        txtno.Text = PIN
        If PIN.StartsWith("P") Then
            query = "select * from  [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] with (nolock) where [P_I_N] = '" & txtno.Text & "' "

        Else
            Label2.Text = "UNIQUE NAV NO:"
            txtno.Text = Nos
            query = "select * from  [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor]with (nolock)  where [No_] = '" & txtno.Text & "' "

        End If
  
         ds = returndataset(query)

        If ds.Tables(0).Rows.Count = 0 Then
            writefailure("Record does not exist in the database")
            btnsave.Visible = False
            Exit Sub
        End If
    End Sub
    Protected Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click

        If rForm.Visible = True Then
            If isvalidfile(fupForm.FileName, False) = False Then
                writefailure("Page 1 is not in the correct format")
                Exit Sub
            End If
            'If duplicateexist(fupForm) = True Then
            '    writefailure("You uploaded same file for Page 1 and one other document " & vbCrLf & "Please re-upload")
            '    Exit Sub
            'End If
        End If

        If r0.Visible = True Then
            If isvalidfile(fupdoc0.FileName, True) = False Then
                writefailure("Page 2 is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupdoc0) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r1.Visible = True Then
            If isvalidfile(fupDoc1.FileName, True) = False Then
                writefailure("Page 3 is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc1) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If



        Try

            System.IO.Directory.CreateDirectory(pathname)
            Dim fullpathstring1 As String = ""
            Dim fullpathstring2 As String = ""
            Dim fullpathstring3 As String = ""
            Dim querycondition As String = ""
            Dim nquery As String = ""
            Dim ds As DataSet
            Dim FPIN As String = ""
            If PIN <> "" Then FPIN = PIN Else FPIN = Nos
            fullpathstring1 = pathname & "\" & FPIN & "-Page 1" & Path.GetExtension(fupForm.FileName)
            fupForm.SaveAs(fullpathstring1)
            ' nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor Rsa Form Path] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','Scanned Form', '" & fullpathstring & "',0,'',0,'','')"
            ' ds = returndataset(nquery)


            '  If r0.Visible Then
            If fupdoc0.FileName <> "" Then
                fullpathstring2 = pathname & "\" & FPIN & "-Page 2" & Path.GetExtension(fupdoc0.FileName)
                fupdoc0.SaveAs(fullpathstring2)
            End If

            If fupDoc1.FileName <> "" Then

                fullpathstring3 = pathname & "\" & FPIN & "-Page 3" & Path.GetExtension(fupDoc1.FileName)
                fupDoc1.SaveAs(fullpathstring3)
            End If

            nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor Rsa Form Path] (Code,[PIN], [Page 1 Url],[Page 2 Url],[Page 3 Url],[Uploaded],[page 1 SP],[page 2 SP],[page 3 SP],[Select],[Selected by])Values('" & Nos & "','" & PIN & "','" & fullpathstring1 & "','" & fullpathstring2 & "','" & fullpathstring3 & "',0,'','','',0,'')"
            ds = returndataset(nquery)
            Response.Write("<script language='javascript'>self.close();</script>")
            writesuccess("Successfully Uploaded, Please close the browser")
        Catch ex As Exception
            writefailure(ex.Message)
        End Try
    End Sub
    Private Sub writefailure(ByVal msg As String)
        lblmsg.Text = msg
    End Sub
    Private Sub writesuccess(ByVal msg As String)
        lblmsg.Text = msg
        lblmsg.ForeColor = Drawing.Color.Green
    End Sub
    Private Function duplicateexist(ByVal fup As FileUpload) As Boolean
        If fup.HasFile = False Then
            Return False
        End If
        If fup.ID <> fupDoc1.ID Then
            If CompareTwoImages(fupDoc1.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
       
        Return False
    End Function
    Private Function isvalidfile(filename As String, ByVal opti As Boolean) As Boolean
        If filename.Trim = "" And opti = True Then Return True
        If filename.Trim = "" Then Return False
        Dim ImgExtension = Path.GetExtension(filename)
        If ImgExtension.ToUpper = ".JPG" Or ImgExtension.ToUpper.Trim = ".JPEG" Or ImgExtension.ToUpper.Trim = ".PDF" Or ImgExtension.ToUpper = ".GIF" Then
            Return True
        Else
            Return False
        End If
    End Function
    Public Function CompareTwoImages(ByVal img1Bytes As Byte(), ByVal img2Bytes As Byte()) As Boolean
        If img1Bytes Is Nothing Or img2Bytes Is Nothing Then
            Return False
        End If

        Dim sha256Managed As New System.Security.Cryptography.SHA256Managed()

        Dim hash1 As Byte() = sha256Managed.ComputeHash(img1Bytes)
        Dim hash2 As Byte() = sha256Managed.ComputeHash(img2Bytes)

        Dim i As Integer = 0
        While i < hash1.Length AndAlso i < hash2.Length
            If hash1(i) = hash2(i) Then
                Return True
            Else
                Return False
            End If
            i += 1
        End While
        Return False
    End Function
    Private Function returndataset(ByVal query As String) As DataSet
        returndataset = Nothing
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            com1 = New SqlCommand(query, con1)
            If con1.State <> ConnectionState.Open Then con1.Open()
            'con1.
            Dim myDataAdapter As New SqlDataAdapter()
            myDataAdapter.SelectCommand = com1
            Dim Request As New DataSet
            myDataAdapter.Fill(Request)
            con1.Close()
            con1.Dispose()
            Return Request
        Catch ex As Exception
            con1.Close()
            con1.Dispose()

            Return Nothing
        End Try
    End Function

End Class
