Imports System.IO
Imports System.Data.SqlClient

Public Class ViewDocumentBen

    Inherits System.Web.UI.Page

    Public docnames As String
    Dim docObj As Object
    Dim pathfiles As String
    Dim lDate As String
    Dim pin As String
    Dim typ As String
    Private Sub writefailure(ByVal msg As String)
        lblmsg.Text = msg
    End Sub
    Private Sub writesuccess(ByVal msg As String)
        lblmsg.Text = msg
        lblmsg.ForeColor = Drawing.Color.Green
    End Sub
    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        pin = Request.QueryString("PIN")
        If pin Is Nothing = False Then
            Dim updatetype As String = Request.QueryString("Typ")
            lblPin.Text = pin
            lblUpdateType.Text = updatetype
            Dim doctype As String

            Dim ds As New DataSet
            Dim query As String
            Dim Nos As String = Request.QueryString("No")
            Dim fullpath As String = ""
            query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Benefit Request Files]  where [Code] ='" & Nos & "'"
            ds = returndataset(query)
            If ds.Tables(0).Rows.Count = 0 Then
                lblmsg.Text = "Attachment not found for this record"
                Exit Sub
            End If
            If ds.Tables(0).Rows.Count > 0 Then
                For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                    fullpath = fullpath & "|" & ds.Tables(0).Rows(i).Item("File Path")
                Next
                fullpath = Mid(fullpath, 2)
            End If

            Dim fullpathobj As Object = Split(fullpath, "|")
            For i As Integer = 0 To UBound(fullpathobj)
                Dim filePath As String = Path.GetDirectoryName(fullpathobj(i).ToString) & "\"
                Dim filename As String = GetFilenameFromPath(fullpathobj(i).ToString)
                Dim filetype As String = Path.GetExtension(fullpathobj(i).ToString)
                If filetype <> "" Then
                    If filetype.ToUpper.Trim = ".PDF" Then
                        doctype = Replace(Path.GetFileNameWithoutExtension(fullpathobj(i).ToString), pin, "")
                        doctype = Replace(doctype, "-", "")
                        continuedFile("FilePage.aspx?FilePath=" & filePath & "&FileName=" & filename & "", doctype)
                    Else
                        doctype = Replace(Path.GetFileNameWithoutExtension(fullpathobj(i).ToString), pin, "")
                        doctype = Replace(doctype, "-", "")
                        continuedImage("ImagePage.aspx?FilePath=" & filePath & "&FileName=" & filename & "", doctype)
                    End If
                End If
            Next
        End If
    End Sub
    Public Function GetFilenameFromPath(FullPath As String) As String
        'Returns "Atomic Kitten - Whole Again.mpg".
        GetFilenameFromPath = Right(FullPath, Len(FullPath) - InStrRev(FullPath, "\"))

    End Function
    Private Sub continuedImage(ByVal imgurl As String, ByVal textstring As String)
        Dim img As New Image
        Dim newDiv As New HtmlGenericControl("div")
        Dim textDiv As New HtmlGenericControl("div")

        ' ADD IMAGE.
        img.ImageUrl = imgurl
        img.Width = "500"
        img.Height = "500"

        newDiv.Attributes.Add("style", "float:left;padding:5px 3px;margin:20px 3px;height:auto;overflow:hidden;")
        newDiv.Controls.Add(img)
        '  newDiv.InnerText = "name"
        ' ADD A TEXT.
        textDiv.Attributes.Add("style", "display:block;font:13px Arial;padding:10px 0;width:img.Width " & img.Width.ToString & "px;color:#666;text-align:center;cursor:pointer;")

        textDiv.InnerText = textstring
        newDiv.Controls.Add(textDiv)
        divGallary.Controls.Add(newDiv)
    End Sub

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
        divGallary.Controls.Add(newDiv)
    End Sub
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