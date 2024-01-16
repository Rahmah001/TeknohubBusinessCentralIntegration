Imports System.IO

Public Class UploadDocuments
    Inherits System.Web.UI.Page
    Public pathname As String
    Public docnames As String
    Dim docObj As Object
    Dim Nos As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        writefailure("")
        If Request.QueryString("Path") <> Nothing Then
            pathname = Request.QueryString("Path")
        End If

        If Request.QueryString("Docs") <> Nothing Then
            docnames = Request.QueryString("Docs")
        End If
        If Request.QueryString("No") <> Nothing Then
            Nos = Request.QueryString("No")
        End If

        docObj = Split(docnames, "|")

        If IsPostBack Then
            '    writebacktofup()
            Exit Sub
        End If

        If Request.QueryString("PIN") <> Nothing Then
            txtno.Text = Request.QueryString("PIN")
        End If

      
        lblarray(docObj)
        Dim db2 As New accountingdDataContext
        Dim venrec As New ARM_FUND_ACCOUNTS_Vendor

        venrec = (From i In db2.ARM_FUND_ACCOUNTS_Vendors Where i.P_I_N = txtno.Text).FirstOrDefault
        If venrec Is Nothing Then
            writefailure("Record does not exist in the database")
            btnsave.Visible = False
            Exit Sub
        End If
    End Sub
    Public Sub lblarray(docname As Object)
        Dim count As Integer
        count = UBound(docname)
        If txtno.Text = "" Then rForm.Visible = False : btnsave.Visible = False
        If count >= 1 Then Label1.Text = docname(0) : r1.Visible = True
        If count >= 2 Then Label2.Text = docname(1) : r2.Visible = True
        If count >= 3 Then Label3.Text = docname(2) : r3.Visible = True
        If count >= 4 Then Label4.Text = docname(3) : r4.Visible = True
    End Sub
    Protected Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click
     
        If rForm.Visible = True Then
            If isvalidfile(fupForm.FileName) = False Then
                writefailure("Scanned Form is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupForm) = True Then
                writefailure("You uploaded same file for Scanned Form and one other document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r1.Visible = True Then
            If isvalidfile(fupdoc1.FileName) = False Then
                writefailure(docObj(0).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupdoc1) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r2.Visible = True Then
            If isvalidfile(fupDoc2.FileName) = False Then
                writefailure(docObj(1).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc2) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        If r3.Visible = True Then
            If isvalidfile(fupDoc3.FileName) = False Then
                writefailure(docObj(2).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc3) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If
        If r4.Visible = True Then
            If isvalidfile(fupDoc4.FileName) = False Then
                writefailure(docObj(3).ToString & "is not in the correct format")
                Exit Sub
            End If
            If duplicateexist(fupDoc4) = True Then
                writefailure("You uploaded same file for more than one document " & vbCrLf & "Please re-upload")
                Exit Sub
            End If
        End If

        Try

            'If fupDoc2.FileBytes.Length > 0 Then rec.SIGNATURE_IMAGE = fupDoc2.FileBytes
            'If fupdoc1.FileBytes.Length > 0 Then rec.PICTURE_IMAGE = fupdoc1.FileBytes
            'If fupDoc3.FileBytes.Length > 0 Then rec.THUMBPRINT = fupDoc3.FileBytes
            'If fupDoc4.FileBytes.Length > 0 Then rec.THUMBPRINT1 = fupDoc4.FileBytes
            ' db.SubmitChanges()  Scanned Form
            System.IO.Directory.CreateDirectory(pathname)
            Dim fullpathstring As String = ""
            Dim querycondition As String = ""

            If rForm.Visible Then
                fullpathstring = pathname & "\" & txtno.Text & "-Scanned Form" & Path.GetExtension(fupForm.FileName)
                fupForm.SaveAs(fullpathstring)
                querycondition = "[Scanned Form] = '" & fullpathstring & "'"
            End If

            If r1.Visible Then
                fullpathstring = pathname & "\" & txtno.Text & "-" & docObj(0).ToString & Path.GetExtension(fupdoc1.FileName)
                fupdoc1.SaveAs(fullpathstring)
                querycondition = querycondition & ", [Document Path 1] = '" & fullpathstring & "'"
            End If

            If r2.Visible Then
                fullpathstring = pathname & "\" & txtno.Text & "-" & docObj(1).ToString & Path.GetExtension(fupDoc2.FileName)
                fupDoc2.SaveAs(fullpathstring)
                querycondition = querycondition & ", [Document Path 2] = '" & fullpathstring & "'"
            End If
            If r3.Visible Then
                fullpathstring = pathname & "\" & txtno.Text & "-" & docObj(2).ToString & Path.GetExtension(fupDoc3.FileName)
                fupDoc3.SaveAs(fullpathstring)
                querycondition = querycondition & ", [Document Path 3] = '" & fullpathstring & "'"
            End If
            If r4.Visible Then
                fullpathstring = pathname & "\" & txtno.Text & "-" & docObj(3).ToString & Path.GetExtension(fupDoc4.FileName)
                fupDoc4.SaveAs(fullpathstring)
                querycondition = querycondition & ", [Document Path 4] = '" & fullpathstring & "'"
            End If
          
            Dim query As String
            Dim conn As New SqlClient.SqlConnection
            Dim comm As New SqlClient.SqlCommand
            query = "update [" & My.MySettings.Default.COMPANYNAME.Trim & "$CSO Rec Update] set " & querycondition & " where PIN= '" & txtno.Text & "' and [No]= '" & Nos & "'"
            conn.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            conn.Open()
            comm = New SqlClient.SqlCommand(query, conn)
            If comm.ExecuteNonQuery = 0 Then
                writefailure("Could not update the document details at this time, Please try again later")
                Exit Sub
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
    Private Function isvalidfile(filename As String) As Boolean
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

    
End Class