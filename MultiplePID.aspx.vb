Imports System.IO
Imports System.Data.SqlClient

Public Class MultiplePID
    Inherits System.Web.UI.Page


    Public pathname As String
    Public docnames As String
    Dim docObj As New ArrayList
    Dim OptionalArray As New ArrayList
    Dim Nos As String
    Dim PIN As String
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        writefailure("")

        If Request.QueryString("nos") <> Nothing Then
            Nos = Request.QueryString("nos")
        End If
        txtno.Text = Nos
        Dim query As String
        If Nos = "" Then Exit Sub
        query = "select * from [Temporary Biometrics]  where TEMPORARY_ID ='" & Nos & "'"
        Dim ds As DataSet = returndataset(query)
        If ds.Tables(0).Rows.Count = 0 Then

            rForm.Visible = False
            writefailure("Invalid record on biometrics table.")

            btnsave.Visible = False
            Exit Sub
        End If
        If PIN Is Nothing Then PIN = ""
    End Sub

    Protected Sub btnsave_Click(sender As Object, e As EventArgs) Handles btnsave.Click

        If rForm.Visible = True Then
            If isvalidfile(fupForm.FileName, False) = False Then
                writefailure("Multiple PID page is not in the correct format")
                Exit Sub
            End If

        End If


        fupForm.SaveAs(Server.MapPath("~/Temp/") + fupForm.FileName)



        Try

            Dim filebyte As Byte() = Nothing
            Dim con As New SqlConnection
            con.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            Dim cmd As SqlCommand = Nothing
            filebyte = System.IO.File.ReadAllBytes(Server.MapPath("~/Temp/") + fupForm.FileName)
            'Dim teststring As String = System.IO.Path.GetFullPath(fupForm.PostedFile.FileName)
            ' filebyte = System.IO.File.ReadAllBytes(System.IO.Path.GetFullPath(fupForm.PostedFile.FileName))
            Dim queries As String = "update [Temporary Biometrics] set [Multiple PID] =@pdf  where TEMPORARY_ID = '" & txtno.Text & "'"
            cmd = New SqlCommand(queries, con)
            cmd.Parameters.Add("@pdf", SqlDbType.Binary).Value = filebyte
            con.Open()
            con.ChangeDatabase("easyRegServer")
            cmd.ExecuteNonQuery()
            con.Close()



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

    Private Function isvalidfile(filename As String, ByVal opti As Boolean) As Boolean
        If filename.Trim = "" And opti = True Then Return True
        If filename.Trim = "" Then Return False
        Dim ImgExtension = Path.GetExtension(filename)
        If ImgExtension.ToUpper.Trim = ".PDF" Then
            Return True
        Else
            Return False
        End If
    End Function
    Private Function returndataset(ByVal query As String) As DataSet
        returndataset = Nothing
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            'con1.ChangeDatabase("easyregServer")
            com1 = New SqlCommand(query, con1)

            If con1.State <> ConnectionState.Open Then con1.Open()
            con1.ChangeDatabase("easyregServer")
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
