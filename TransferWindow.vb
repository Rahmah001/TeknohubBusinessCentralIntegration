Imports System.IO
Imports System.Net
Imports System.Web.Script.Serialization
Imports System.Xml.Serialization
Imports RestSharp
Public Class TransferWindow
    Public Function SubmitProvisionalBalance(ByVal userid As String, pwd As String, tok As String, qtr As String) As String

        'Dim Clients As New WebClient
        Dim Requests As HttpWebRequest = HttpWebRequest.Create("Http://rts.pencom.gov.ng/submitprovbal")
        Dim Responses As HttpWebResponse
        Dim Distances As String
        'Requests = New WebRequest
        ' Clients.BaseAddress = "Https://rts.pencom.gov.ng/submitprovbal"
        Requests.Credentials = CredentialCache.DefaultCredentials
        Requests.Method = "POST"
        Requests.ContentType = "application/json"
        Requests.Headers.Add(System.Net.HttpRequestHeader.Cookie, "security=true")
        Requests.AllowAutoRedirect = False
        'Requests.Headers.Add("Content-Type", "application/json") '*** The line that made the code work On mac
        Dim data As String = "userid=test&password=test&token=test&quarter=test&transferRefId =test" ' //replace <value>
        Dim dataStream As Byte() = Encoding.UTF8.GetBytes(data)
        Requests.ContentLength = dataStream.Length
        Dim newStream As Stream = Requests.GetRequestStream()

        newStream.Write(dataStream, 0, dataStream.Length)
        newStream.Close()
        Responses = Requests.GetResponse
        'Requests.AddParameter("userid", "test")
        'request.AddParameter("password", "test")
        'request.AddParameter("token", "test")
        'request.AddParameter("quarter", "test")
        'request.AddParameter("transferRefId", "test")



        'Responses = Clients.Execute(Requests,)

        'response = Clients.Execute(Requests)

        Return Responses.StatusCode.ToString + " " + Responses.StatusDescription.ToString

    End Function

End Class
