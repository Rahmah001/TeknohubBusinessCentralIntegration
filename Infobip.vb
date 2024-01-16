Imports RestSharp
Imports System.Web.Script.Serialization
Imports System.Net
Imports RestSharp.Authenticators

Public Class Infobip

    Private infourl As RestClient = New RestClient("https://api.infobip.com/sms/1/text/single")
    Private requestTimeStamp As String = ""

    Dim API_KEY As String = "jkgud897jgkdkkd87t"
    Public Function sendsmsInfobip(ByVal phoneno As String, ByVal msg As String) As String
        sendsmsInfobip = ""
        Dim model As New SendProperty
        model.from = "ARMPENSIONS"
        model.to = phoneno
        model.text = msg

        Dim request = New RestRequest("", Method.POST)
        request.AddHeader("Content-Type", "application/json; charset=utf-8")
        infourl.Authenticator = New HttpBasicAuthenticator("ARMPen1", "Test1234")

        request.AddJsonBody(model)

        Dim Response As IRestResponse = infourl.Execute(request)



        If Response.IsSuccessful = False Then
            MsgBox(Response.StatusDescription)

            Return Response.StatusDescription
        Else
            Return "Success"
        End If
    End Function



End Class
