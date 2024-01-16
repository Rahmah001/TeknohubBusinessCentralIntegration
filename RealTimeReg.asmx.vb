Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://armpension.com/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class RealTimeReg
    Inherits System.Web.Services.WebService
    <WebMethod()> _
    Public Function GetNewClient(ByVal orderid As String, ByVal filepath As String) As String
        Dim dt As New DataTable
        Dim ds As New DataSet

        Try
            If filepath.Trim <> "" Then
                ds.ReadXml(filepath)
            End If
            ' If ds Is Nothing Then ds.ReadXml(New StringReader(xmlstring))
        Catch ex As Exception
            Return "INVALID XML FORMAT" & vbCrLf & ex.Message
        End Try

        Try
            dt = ds.Tables("request")
        Catch ex As Exception
            Return "Unable to read the REQUEST table" & vbCrLf & ex.Message
        End Try
        Try
            Dim dlltouse As New MovefromAbby14.GetXmltoVendor
            GetNewClient = dlltouse.insertrecordtovendor(orderid, dt)

            Return GetNewClient
        Catch ex As Exception
            '2    Return returnxmlstring("", "F", ex.Message)
            GetNewClient = ex.Message
        End Try
        Return GetNewClient
    End Function
   
End Class