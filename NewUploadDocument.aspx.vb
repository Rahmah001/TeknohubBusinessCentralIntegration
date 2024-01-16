Imports System.IO
Imports System.Data.SqlClient

Public Class NewUploadDocument
    Inherits System.Web.UI.Page


    Public pathname As String
    Public docnames As String
    Dim docObj As New ArrayList
    Dim OptionalArray As New ArrayList
    Dim Nos As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        writefailure("")
        If Request.QueryString("Path") <> Nothing Then
            pathname = Request.QueryString("Path")
        End If

       
        If Request.QueryString("No") <> Nothing Then
            Nos = Request.QueryString("No")
        End If


        If Request.QueryString("No") = Nothing Then
            Exit Sub
        End If
        Dim query As String

        query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File]  where [Code] ='" & Nos & "'"
        Dim ds As DataSet = returndataset(query)
        If ds.Tables(0).Rows.Count > 0 Then
            rForm.Visible = False
            writefailure("Documents previously attached for this record.")
            btnsave.Visible = False
            Exit Sub
        End If


        query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$CSO Rec Update Line] where [No] ='" & Nos & "'"
        ds = returndataset(query)
        If ds Is Nothing Then
            rForm.Visible = False
            writefailure("Record doesn't require any update")
            btnsave.Visible = False
            Exit Sub
        End If
        Dim reqtypel As String = ""
        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
            reqtypel = reqtypel & " [Req Update Type] = '" & ds.Tables(0).Rows(i).Item("Update Type").ToString & "' or "
        Next
        reqtypel = Mid(reqtypel, 1, reqtypel.Trim.Length - 2)

        query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Req Upd Document List] where " & reqtypel

        ds = returndataset(query)

        For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
            Dim flname As String = ds.Tables(0).Rows(i).Item("Document Name")
            If Not (docObj.Contains(flname)) Then
                docObj.Add(flname)
                OptionalArray.Add(ds.Tables(0).Rows(i).Item("Optional"))
            End If
        Next

        If IsPostBack Then
            Exit Sub
        End If

        If Request.QueryString("PIN") <> Nothing Then
            txtno.Text = Request.QueryString("PIN")
        End If


        lblarray(docObj)
        query = "select * from  [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] where [P_I_N] = '" & txtno.Text & "' "
        ds = returndataset(query)

        If ds.Tables(0).Rows.Count = 0 Then
            writefailure("Record does not exist in the database")
            btnsave.Visible = False
            Exit Sub
        End If
    End Sub
    Public Sub lblarray(docname As ArrayList)
        Dim count As Integer
        count = (docname).Count
        If txtno.Text = "" Then rForm.Visible = False : btnsave.Visible = False
        If count >= 1 Then Label0.Text = docname(0) : r0.Visible = True
        If count >= 2 Then Label1.Text = docname(1) : r1.Visible = True
        If count >= 3 Then Label2.Text = docname(2) : r2.Visible = True
        If count >= 4 Then Label3.Text = docname(3) : r3.Visible = True
        If count >= 5 Then Label4.Text = docname(4) : r4.Visible = True
        If count >= 6 Then Label5.Text = docname(5) : r5.Visible = True
        If count >= 7 Then Label6.Text = docname(6) : r6.Visible = True
        If count >= 8 Then Label7.Text = docname(7) : r7.Visible = True
        If count >= 9 Then Label8.Text = docname(8) : r8.Visible = True
        If count >= 10 Then Label9.Text = docname(9) : r9.Visible = True

    End Sub
    Protected Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click

        If rForm.Visible = True Then
            If isvalidfile(fupForm.FileName, False) = False Then
                writefailure("Scanned Form is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupForm) = True Then
                writefailure("You uploaded same file for Scanned Form and one other document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r0.Visible = True Then
            If isvalidfile(fupdoc0.FileName, OptionalArray(0)) = False Then
                writefailure(docObj(0).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupdoc0) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r1.Visible = True Then
            If isvalidfile(fupDoc1.FileName, OptionalArray(1)) = False Then
                writefailure(docObj(1).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc1) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r2.Visible = True Then
            If isvalidfile(fupDoc2.FileName, OptionalArray(2)) = False Then
                writefailure(docObj(2).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc2) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r3.Visible = True Then
            If isvalidfile(fupDoc3.FileName, OptionalArray(3)) = False Then
                writefailure(docObj(3).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc3) = True Then

                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r4.Visible = True Then
            If isvalidfile(fupDoc4.FileName, OptionalArray(4)) = False Then
                writefailure(docObj(4).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc4) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r5.Visible = True Then
            If isvalidfile(fupDoc5.FileName, OptionalArray(5)) = False Then
                writefailure(docObj(5).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc5) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r6.Visible = True Then
            If isvalidfile(fupDoc6.FileName, OptionalArray(6)) = False Then
                writefailure(docObj(6).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc6) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r7.Visible = True Then
            If isvalidfile(fupDoc7.FileName, OptionalArray(7)) = False Then
                writefailure(docObj(7).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc7) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r8.Visible = True Then
            If isvalidfile(fupDoc8.FileName, OptionalArray(8)) = False Then
                writefailure(docObj(8).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc8) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r9.Visible = True Then
            If isvalidfile(fupDoc9.FileName, OptionalArray(9)) = False Then
                writefailure(docObj(9).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc9) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If


        Try

            System.IO.Directory.CreateDirectory(pathname)
            Dim fullpathstring As String = ""
            Dim querycondition As String = ""
            Dim nquery As String = ""
            Dim ds As DataSet
            If rForm.Visible Then
                fullpathstring = pathname & "\" & Nos & "-Scanned Form" & Path.GetExtension(fupForm.FileName)
                fupForm.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','Scanned Form', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If

            If r0.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(0).ToString & Path.GetExtension(fupdoc0.FileName)
                fupdoc0.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(0).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If

            If r1.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(1).ToString & Path.GetExtension(fupDoc1.FileName)
                fupDoc1.SaveAs(fullpathstring)
                '   querycondition = querycondition & ", [Document Path 2] = '" & fullpathstring & "'"
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(1).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r2.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(2).ToString & Path.GetExtension(fupDoc2.FileName)
                fupDoc2.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(2).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r3.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(3).ToString & Path.GetExtension(fupDoc3.FileName)
                fupDoc3.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(3).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r4.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(4).ToString & Path.GetExtension(fupDoc4.FileName)
                fupDoc4.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(4).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If

            If r5.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(5).ToString & Path.GetExtension(fupDoc5.FileName)
                fupDoc5.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(5).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r6.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(6).ToString & Path.GetExtension(fupDoc6.FileName)
                fupDoc6.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(6).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r7.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(7).ToString & Path.GetExtension(fupDoc7.FileName)
                fupDoc7.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(7).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r8.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(8).ToString & Path.GetExtension(fupDoc8.FileName)
                fupDoc8.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(8).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
            If r9.Visible Then
                fullpathstring = pathname & "\" & Nos & "-" & docObj(9).ToString & Path.GetExtension(fupDoc9.FileName)
                fupDoc9.SaveAs(fullpathstring)
                nquery = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Rec_ Update File] (Code, [Document Name],[File Path],[Select],[Original File Path],[Uploaded],[Response Message],[Selected By]) Values('" & Nos & "','" & docObj(9).ToString & "', '" & fullpathstring & "',0,'',0,'','')"
                ds = returndataset(nquery)
            End If
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
        If fup.ID <> fupdoc1.ID Then
            If CompareTwoImages(fupdoc1.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupDoc4.ID Then
            If CompareTwoImages(fupDoc4.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupDoc3.ID Then
            If CompareTwoImages(fupDoc3.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupDoc2.ID Then
            If CompareTwoImages(fupDoc2.FileBytes, fup.FileBytes) = True Then
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
