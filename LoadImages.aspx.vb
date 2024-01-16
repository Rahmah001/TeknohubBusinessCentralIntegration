Imports System.IO

Public Class LoadImages
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        writefailure("")
        If IsPostBack Then Exit Sub
        If Request.QueryString("NO") <> Nothing Then
            txtno.Text = Request.QueryString("NO")
        End If
        Dim db As New PersonalDetailsDataContext
        Dim rec As New empImage
        Dim db2 As New accountingdDataContext
        Dim venrec As New ARM_FUND_ACCOUNTS_Vendor


        venrec = (From i In db2.ARM_FUND_ACCOUNTS_Vendors Where i.No_ = txtno.Text).FirstOrDefault
        If venrec Is Nothing Then
            writefailure("Record does not exist in the database")
            btnsave.Visible = False
            Exit Sub
        Else
            rec = (From i In db.empImages Where i.REGISTRATION_CODE_N = txtno.Text).FirstOrDefault
            If rec Is Nothing Then
                ' writefailure("Record does not exist in the images database")
                'btnsave.Visible = False
                'Exit Sub
            End If
        End If
        If venrec.Temp = False Then
            writefailure("Record already approved")
            btnsave.Visible = False
            Exit Sub
        End If

        If venrec.Status <> 0 And venrec.Status <> 2 Then
            writefailure("Record already has biometrics details and cannot be edited")
            btnsave.Visible = False
            Exit Sub
        End If
        Session("exist") = False
        If rec Is Nothing = False Then
            If Not (rec.SIGNATURE_IMAGE Is Nothing) Then
                Session("exist") = True
                ' writefailure("The newly uploaded biometrics will overwrite the previous one there" & vbCrLf & "Please close the browser to discard update")
                writefailure("This record already has biometrics data, To update the existing data, browse and upload " & vbCrLf & " Otherwise, close the browser to discard update")

                Exit Sub
            End If
        End If
    End Sub

    Protected Sub txtsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click

        If Session("exist") = False Then
            If fupleftthumb.HasFile = False Or fuppassport.HasFile = False Or fupRightthumb.HasFile = False Or fupSignature.HasFile = False Then
                writefailure("One of the biometrics details is yet to be uploaded." + vbCrLf + "Please correct and retry")
                Exit Sub
            End If
        End If
        If isvalidfile(fupleftthumb.FileName) = False Or isvalidfile(fupRightthumb.FileName) = False Then
            writefailure("One or more of the fingerprint is not in the correct format")
            Exit Sub
        End If
        If isvalidfile(fuppassport.FileName) = False Or isvalidfile(fupSignature.FileName) = False Then
            writefailure("Either the passport or the signature is not in the correct format")
            Exit Sub
        End If

        If duplicateexist(fupleftthumb) = True Then
            writefailure("You uploaded same file for more than one biometrics details" & vbCrLf & "Please re-upload")
            Exit Sub
        End If
        If duplicateexist(fuppassport) = True Then
            writefailure("You uploaded same file for more than one biometrics details" & vbCrLf & "Please re-upload")
            Exit Sub
        End If
        If duplicateexist(fupRightthumb) = True Then
            writefailure("You uploaded same file for more than one biometrics details" & vbCrLf & "Please re-upload")
            Exit Sub
        End If
        If duplicateexist(fupSignature) = True Then
            writefailure("You uploaded same file for more than one biometrics details" & vbCrLf & "Please re-upload")
            Exit Sub
        End If
        Dim isinimagestable As Boolean = True
        Try
            Dim db As New PersonalDetailsDataContext
            Dim rec As New empImage
            rec = (From i In db.empImages Where i.REGISTRATION_CODE_N = txtno.Text).FirstOrDefault


            If rec Is Nothing Then
                ' writefailure("Record does not exist in the database")
                'Exit Sub
                rec = New empImage
                rec.REGISTRATION_CODE = txtno.Text
                rec.REGISTRATION_CODE_N = txtno.Text
                isinimagestable = False
            End If
            'If rec.SIGNATURE_IMAGE <> Nothing Then
            '    writefailure("Record already has biometrics details and cannot be edited")
            '    Exit Sub
            'End If
            If fupSignature.FileBytes.Length > 0 Then rec.SIGNATURE_IMAGE = fupSignature.FileBytes
            If fuppassport.FileBytes.Length > 0 Then rec.PICTURE_IMAGE = fuppassport.FileBytes
            If fupRightthumb.FileBytes.Length > 0 Then rec.THUMBPRINT = fupRightthumb.FileBytes
            If fupleftthumb.FileBytes.Length > 0 Then rec.THUMBPRINT1 = fupleftthumb.FileBytes
            If isinimagestable = False Then db.empImages.InsertOnSubmit(rec)
            db.CommandTimeout = 0
            db.SubmitChanges()

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
        If fup.ID <> fuppassport.ID Then
            If CompareTwoImages(fuppassport.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupleftthumb.ID Then
            If CompareTwoImages(fupleftthumb.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupRightthumb.ID Then
            If CompareTwoImages(fupRightthumb.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        If fup.ID <> fupSignature.ID Then
            If CompareTwoImages(fupSignature.FileBytes, fup.FileBytes) = True Then
                Return True
            End If
        End If
        Return False
    End Function
    Private Function isvalidfile(filename As String) As Boolean
        If filename.Trim = "" Then Return True
        Dim ImgExtension = Path.GetExtension(filename)
        If ImgExtension.ToUpper = ".JPG" Or ImgExtension.ToUpper.Trim = ".JPEG" Then
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