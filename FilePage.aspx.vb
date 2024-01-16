Imports System.IO
Public Class FilePage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Request.QueryString("FileName") IsNot Nothing Then
            Try

                ' Read the file and convert it to Byte Array
                '"\\S-HQP-DMGT01\UploadedDocs\"
                Dim filePath As String = Request.QueryString("FilePath")

                Dim filename As String = Request.QueryString("FileName")
                ' Dim filename As String = "PEN100689851821-Birth Certificate.jpg"
                Dim contenttype As String = "application/pdf"

                Dim fs As FileStream = New FileStream(filePath & filename, FileMode.Open, FileAccess.Read)
                Dim br As BinaryReader = New BinaryReader(fs)
                Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))
                br.Close()
                fs.Close()

                'Write the file to Reponse
                Response.Buffer = True
                Response.Charset = ""
                '   Context.Response.AppendHeader("Content-Disposition", Convert.ToString("attachment; filename=") & filename)
                Response.Cache.SetCacheability(HttpCacheability.NoCache)
                Response.ContentType = contenttype
                '   Response.AddHeader("content-disposition", "attachment;filename=" & filename)
                Response.BinaryWrite(bytes)
                Response.Flush()
                Response.End()
            Catch ex As Exception
                Write(ex.Message)
            End Try


        End If
    End Sub

End Class