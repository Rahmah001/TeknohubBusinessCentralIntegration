Imports System.Data.SqlClient

Public Class ViewPID
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim no = Request.QueryString("Nos")
        If no Is Nothing Then

            Exit Sub
        End If
        Dim updatetype As String = " PID"


        Dim ds As New DataSet
            Dim query As String
            Dim Nos As String = Request.QueryString("No")
            Dim fullpath As String = ""
            query = "select [Multiple PID]  from [Temporary Biometrics] where TEMPORARY_ID='" & no & "'"
            ds = returndataset(query)
            If ds.Tables(0).Rows.Count = 0 Then
                lblmsg.Text = "Attachment not found for this record"
                Exit Sub
            End If

        If IsDBNull(ds.Tables(0).Rows(0).Item("Multiple PID")) Then
            lblmsg.Text = "No Attachment found for this record"
            Exit Sub
        End If

        Dim byteval As Byte() = ds.Tables(0).Rows(0).Item("Multiple PID")

        Response.Buffer = True
            Response.Charset = " "
            Response.Cache.SetCacheability(HttpCacheability.NoCache)
            Response.ContentType = "application/pdf"
            Response.BinaryWrite(byteval)
            Response.Flush()
            Response.End()




    End Sub
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
    Private Sub continuedFile(ByVal fileurl As String, ByVal textstring As String)
        Dim newDiv As New HtmlGenericControl("div")
        Dim textDiv As New HtmlGenericControl("div")

        ' ADD IMAGE.

        Dim ifram As New HtmlGenericControl("iframe")
        ' ifram.Attributes.Add("src", "http://docs.google.com/gview?url=http://infolab.stanford.edu/pub/papers/google.pdf&embedded;=true")
        ifram.Attributes.Add("src", fileurl.Trim & "&embedded;=true")
        ifram.Attributes.Add("style", "width:600px; height:500px;")
        ifram.Attributes.Add("frameborder", "5")

        newDiv.Attributes.Add("style", "float:left;padding:5px 3px;margin:20px 3px;height:auto;overflow:hidden;")
        newDiv.Controls.Add(ifram)

        ' ADD A TEXT.
        textDiv.Attributes.Add("style", "display:block;font:13px Arial;padding:10px 0;width:img.Width " & 500 & "px;color:#666;text-align:center;cursor:pointer;")

        textDiv.InnerText = textstring
        newDiv.Controls.Add(textDiv)
        'DIvGallary.Controls.Add(newDiv)
    End Sub
End Class