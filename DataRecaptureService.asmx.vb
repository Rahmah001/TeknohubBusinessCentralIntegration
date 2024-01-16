Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports RestSharp

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class DataRecaptureService
    Inherits System.Web.Services.WebService
    Dim errorresp As String = ""
    <WebMethod()>
    Public Function GetNINConsent(ByVal code As String) As String
        Dim query As String
        Dim dst As New DataSet
        Dim errormsg As String = ""
        Dim easyconstring As String = "easyregserver"
        Try
            query = "Select * from [enrollee] where code = '" & code & "'"
            dst = returndataset(easyconstring, query)

            If dst.Tables(0).Rows.Count = 0 Then

                Return Nothing
            End If

            If dst.Tables(0).Rows(0).Item("registration_type") = "CFI" Then
                Dim rrcode As String = ""
                If IsDBNull(dst.Tables(0).Rows(0).Item("Recapture_Code")) Then
                    dst.Tables(0).Rows(0).Item("Recapture_Code") = 0
                End If


                If dst.Tables(0).Rows(0).Item("Recapture_Code") = 0 Then
                    Dim recapcode As String = ("select top 1 Recapture_Code from enrollee where Recapture_Code <> '' order by Recapture_Code desc")
                    Dim recapds As New DataSet

                    Dim lastnum As Integer
                    recapds = returndataset(easyconstring, recapcode)
                    If recapds.Tables(0).Rows.Count > 0 Then
                        lastnum = recapds.Tables(0).Rows(0).Item("Recapture_Code") + 1
                    Else
                        lastnum = 1
                    End If
                    rrcode = "update enrollee set Recapture_Code = '" & lastnum & "' where [code] = '" & code & "' "
                    recapds = returndataset(easyconstring, rrcode)
                End If
                Dim ninurl As String = My.MySettings.Default.ninurl + "cfi/" + code
                Dim infourl As RestClient = New RestClient(ninurl)
                Dim request = New RestRequest(ninurl, Method.GET)
                request.AddHeader("Content-Type", "application/json; charset=utf-8")
                'request.AddJsonBody(model)
                Dim Response As IRestResponse = infourl.Execute(request)
                If Response.IsSuccessful = False Then
                    Return Response.ErrorMessage
                Else
                    Return Response.StatusDescription
                End If
                Return Response.StatusDescription
            End If
            If dst.Tables(0).Rows(0).Item("registration_type") = "RSA" Then
                Dim ninurl As String = My.MySettings.Default.ninurl + "rsa/" + code
                Dim infourl As RestClient = New RestClient(ninurl)
                Dim request = New RestRequest(ninurl, Method.GET)
                request.AddHeader("Content-Type", "application/json; charset=utf-8")
                'request.AddJsonBody(model)
                Dim Response As IRestResponse = infourl.Execute(request)
                Return Response.StatusDescription
            End If


        Catch ex As Exception
            Return ex.Message

        End Try
        Return "nothing"
    End Function
    Public Function vstring(ByVal values As String) As String
        Dim begining As String = Mid(values, 1, 22)
        If begining.EndsWith(",") Then
            Return Mid(values, 23)
        Else
            Return Mid(values, 24)
        End If

    End Function
    Public Function stringToBase64ByteArray(ByVal input As String) As Byte()

        Try
            Dim ret As Byte() = System.Text.ASCIIEncoding.Default.GetBytes(input)
            '  Dim s As String = System.Convert.ToBase64String(ret)
            Dim s As String = System.Text.Encoding.UTF8.GetString(ret)
            ' Dim s As String = Convert.ToBase64String(ret, Base64FormattingOptions.InsertLineBreaks)
            ret = System.Convert.FromBase64String(s)
            Return ret
        Catch EX As Exception
            Return Nothing
        End Try
    End Function
    Private Function returndataset(ByVal databasenames As String, ByVal query As String) As DataSet
        errorresp = ""
        returndataset = Nothing
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            'con1.ConnectionString = "Data Source=192.168.0.30;Initial Catalog=ARM;User ID=sa;Password=*aiico1"

            con1.Open()
            If databasenames <> "" Then con1.ChangeDatabase(databasenames)
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

            errorresp = ex.Message
            Return Nothing
        End Try
    End Function
End Class