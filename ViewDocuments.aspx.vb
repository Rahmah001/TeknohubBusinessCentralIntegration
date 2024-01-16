Imports System.IO
Public Class ViewDocuments

    Inherits System.Web.UI.Page

    Public docnames As String
    Dim docObj As Object
    Dim pathfiles As String
    Dim lDate As String
    Dim pin As String
    Dim typ As String

    'Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    Dim iFileCnt As Integer = 0
    '    If Request.QueryString("PathFiles") <> Nothing Then
    '        pathfiles = Request.QueryString("PathFiles")
    '    End If
    '    docObj = Split(pathfiles, "|")

    '    If Request.QueryString("Docs") <> Nothing Then
    '        docnames = Request.QueryString("Docs")
    '    End If
    '    Dim dirInfo As New System.IO.DirectoryInfo("\\S-HQP-DMGT01\UploadedDocs\")
    '    Dim listfiles As FileInfo() = dirInfo.GetFiles("*.*")

    '    If listfiles.Length > 0 Then
    '        For Each file As FileInfo In listfiles

    '             CHECK THE TYPE OF FILE.
    '            If Trim(listfiles(iFileCnt).Extension) = ".jpg" Or _
    '                Trim(listfiles(iFileCnt).Extension) = ".jpeg" Or _
    '                    Trim(listfiles(iFileCnt).Extension) = ".png" Or _
    '                        Trim(listfiles(iFileCnt).Extension) = ".bmp" Or _
    '                            Trim(listfiles(iFileCnt).Extension) = ".gif" Then

    '                Dim img As New HtmlImage
    '                Dim newDiv As New HtmlGenericControl("div")
    '                Dim textDiv As New HtmlGenericControl("div")

    '                ADD Image.
    '               img.Src = "\\S-HQP-DMGT01\UploadedDocs\" & listfiles(iFileCnt).Name
    '                img.Src = Server.("\\S-HQP-DMGT01\UploadedDocs\" & listfiles(iFileCnt).Name)
    '                img.Src = "C:/NAV Licence/test/PEN100689851821-Birth Certificate.jpg"
    '                img.Width = "130"
    '                img.Height = "130"

    '                newDiv.Attributes.Add("style", "float:left;padding:5px 3px;margin:20px 3px;height:auto;overflow:hidden;")
    '                newDiv.Controls.Add(img)

    '                 ADD A TEXT.
    '                textDiv.Attributes.Add("style", "display:block;font:13px Arial;padding:10px 0;width:" & _
    '                    img.Width & "px;color:#666;text-align:center;cursor:pointer;")

    '                textDiv.InnerText = "Various Catgories - Binding, Product Specification, Author (Price Tag)"
    '                newDiv.Controls.Add(textDiv)

    '                divGallary.Controls.Add(newDiv)

    '            End If

    '            iFileCnt = iFileCnt + 1

    '        Next
    '    End If
    'End Sub
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
            Dim fullpath As String = Request.QueryString("FullPath")
            '     Dim ds As New DataSet
            '    Dim query As String
            '   query = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Benefit Request Files]  where [Code] ='"&  &"'"



            Dim fullpathobj As Object = Split(fullpath, "|")
            For i As Integer = 0 To UBound(fullpathobj)
                Try

               
                Dim filePath As String = Path.GetDirectoryName(fullpathobj(i).ToString) & "\"
                Dim filename As String = GetFilenameFromPath(fullpathobj(i).ToString)
                Dim filetype As String = Path.GetExtension(fullpathobj(i).ToString)
                ' Image1.ImageUrl = "ImageVB.aspx?FilePath=" & filePath & "&FileName=" & filename & ""
                If filetype.ToUpper.Trim = ".PDF" Then
                    doctype = Replace(Path.GetFileNameWithoutExtension(fullpathobj(i).ToString), pin, "")
                    doctype = Replace(doctype, "-", "")
                    continuedFile("FilePage.aspx?FilePath=" & filePath & "&FileName=" & filename & "", doctype)
                Else
                    doctype = Replace(Path.GetFileNameWithoutExtension(fullpathobj(i).ToString), pin, "")
                    doctype = Replace(doctype, "-", "")
                    continuedImage("ImagePage.aspx?FilePath=" & filePath & "&FileName=" & filename & "", doctype)
                End If
                Catch ex As Exception

                End Try
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
End Class