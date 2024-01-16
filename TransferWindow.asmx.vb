Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlTypes
Imports System.Drawing
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Security.Cryptography.X509Certificates
Imports System.Net.Security
Imports RestSharp
Imports RestSharp.Authenticators
Imports WebApplication1.QuickType
Imports Newtonsoft.Json
Imports Newtonsoft
Imports System.Xml.Serialization
Imports System.Xml
Imports AiicoAssistance.QuickType


' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class TransferWindow1
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function SubmitProvBalances(ByVal userid As String, pwd As String, tok As String, qtr As String, ByVal refid As String) As String
        'Dim infourl As RestClient = New RestClient("Http://rts.pencom.gov.ng/submitprovbal")
        If userid = "" Or pwd = "" Or tok = "" Or qtr = "" Or refid = "" Then
            Return ("Error: Invalid Credentials")
        End If
        Dim refobj As Object = Split(refid, "|")
        Dim refstrings As String = ""
        Dim ds As New DataSet
        For i As Integer = 0 To UBound(refobj)
            refstrings = refstrings & ",'" & refobj(i).ToString & "'"
        Next
        refstrings = Mid(refstrings, 2)
        If qtr.Length <> 6 Then Return "Error: Invalid Quarter"
        ' Dim eftsharedservice As New EFTIntegration
        Dim isupdate As Boolean = True
        Dim usera = New ProvBalheader
        Dim recBody = New ProvbalBody
        usera.userid = userid
        With usera
            .userid = userid
            .password = pwd
            .quarter = qtr
            .token = tok

        End With
        Dim sqlquery As String = ""
        Try
            sqlquery = "Select * from  [ARM FUND ACCOUNTS$Projected RSA Balance] where  TransferRef in ( " & refstrings & ")"
            ds = returndataset(sqlquery)

        Catch ex As Exception
            Return ex.Message & " " & sqlquery
        End Try
        Dim dt As New DataTable
        If ds Is Nothing Then
            Return "Error: Cannot read from Projected RSA Balance with refids " & refstrings
        End If

        If ds.Tables(0).Rows.Count = 0 Then
            Return "Error: Cannot read from Projected RSA Balance with refids " & refstrings & sqlquery
        End If

        dt = ds.Tables(0)

        dt.AsEnumerable().Where(Function(row) row.ItemArray.All(Function(field) field Is Nothing Or field Is DBNull.Value Or field.Equals(""))).ToList().ForEach(Sub(row) row.Delete())
        dt.AcceptChanges()
        Try
            Dim robject As New ProvObject


            Dim provisionalDetail As New List(Of provisionalDetail)
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim trans As New provisionalDetail
                'trans.referenceId = "043-180920-0020"
                'trans.asAtDate = "18-Sep-2020"
                'trans.rsaBalance = 833285.52

                trans.referenceId = dt.Rows(i).Item(1).ToString
                trans.asAtDate = dt.Rows(i).Item(4).ToString
                trans.rsaBalance = Math.Round(dt.Rows(i).Item(2), 4)
                provisionalDetail.Add(trans)
            Next
            robject.provisionalDetails = provisionalDetail

            Dim hp As New RestSharp.Helper.ApiHelper
            Dim ret As String = hp.SubmitBalances(My.MySettings.Default.TWEndPoint, "/submitprovbal", usera, robject)



            If ret.ToUpper.StartsWith("SUCCESS") = False Then Return ret

        Catch ex As Exception
            Return ex.Message & "--" & dt.Rows.Count
        End Try
        Return "success"
    End Function

    <WebMethod()>
    Public Function DownloadProvBalances(ByVal userid As String, pwd As String, tok As String, qtr As String, ByVal refid As String) As String
        If userid = "" Or pwd = "" Or tok = "" Or qtr = "" Then
            Return ("Error: Invalid Credentials")
        End If
        Dim refobj As Object = Split(refid, "|")
        Dim refstrings As String = ""
        Dim ds As New DataSet
        For i As Integer = 0 To UBound(refobj)
            refstrings = refstrings & ",'" & refobj(i).ToString & "'"
        Next
        refstrings = Mid(refstrings, 2)
        If qtr.Length <> 6 Then Return "Error: Invalid Quarter"

        'Dim eftsharedservice As New EFTIntegration
        Dim isupdate As Boolean = True
        Dim usera = New ProvBalheader
        Dim recBody = New ProvbalBody
        usera.userid = userid
        With usera
            .userid = userid
            .password = pwd
            .quarter = qtr
            .token = tok
        End With
        Dim hp As New RestSharp.Helper.ApiHelper
        Dim ret As String = hp.DownloadBalances(My.MySettings.Default.TWEndPoint, "/downloadprovbal", usera)


        If ret.ToUpper.StartsWith("SUCCESS") = False Then Return ret

        'Catch ex As Exception
        '    Return ex.Message
        'End Try
        Return "success"
    End Function

    <WebMethod()>
    Public Function SubmitTransferHistory(ByVal userid As String, pwd As String, tok As String, qtr As String, ByVal refid As String, ByVal xmlfilepath As String) As String
        'Dim infourl As RestClient = New RestClient("Http://rts.pencom.gov.ng/submitth")
        If userid = "" Or pwd = "" Or tok = "" Or qtr = "" Or refid = "" Or xmlfilepath = "" Then
            Return ("Error: Invalid/Empty Credentials")
        End If
        If qtr.Length <> 6 Then Return "Error: Invalid Quarter"

        If xmlfilepath = "" Then xmlfilepath = "c:\temp\PEN200856990620TransferOut.xml"
        Dim ds As New DataSet

        Dim robject As New thObject
        'Dim eftsharedservice As New EFTIntegration
        Dim isupdate As Boolean = True
        Dim usera = New ProvBalheader
        Dim recBody = New ProvbalBody
        usera.userid = userid
        With usera
            .userid = userid
            .password = pwd
            .quarter = qtr
            .token = tok

        End With

        Dim thds As New DataSet
        'thds.ReadXml(xmlfilepath)
        Dim thdetails As New List(Of detailRecord)
        convertxmlfiletods(xmlfilepath, qtr, refid, robject)
        'Return thds.Tables.Count
        'Dim sqlquery As String = ""
        Try
            Dim hp As New RestSharp.Helper.ApiHelper

            Dim ret As String = hp.SubmitTH(My.MySettings.Default.TWEndPoint, "/submitth", refid, usera, robject)



            If ret.ToUpper.StartsWith("SUCCESS") = False Then Return ret
            '            Return ret

        Catch ex As Exception
            Return ex.Message
        End Try
        Return "success"
    End Function


    <WebMethod()>
    Public Function DownloadTHistory(ByVal userid As String, pwd As String, tok As String, qtr As String, ByVal refid As String) As String
        If userid = "" Or pwd = "" Or tok = "" Or qtr = "" Then
            Return ("Error: Invalid Credentials")
        End If
        Dim refobj As Object = Split(refid, "|")
        Dim refstrings As String = ""
        Dim ds As New DataSet
        If qtr.Length <> 6 Then Return "Error: Invalid Quarter"

        'Dim eftsharedservice As New EFTIntegration
        Dim isupdate As Boolean = True
        Dim usera = New ProvBalheader
        Dim recBody = New ProvbalBody
        usera.userid = userid
        With usera
            .userid = userid
            .password = pwd
            .quarter = qtr
            .token = tok

        End With
        Dim hp As New RestSharp.Helper.ApiHelper
        Dim ret As String = hp.DownloadTH(My.MySettings.Default.TWEndPoint, "/downloadth", usera, refid)


        If ret.ToUpper.StartsWith("SUCCESS") = False Then Return ret

        'Catch ex As Exception
        '    Return ex.Message
        'End Try
        Return "success"
    End Function
    Private Function convertxmlfiletods(ByVal filenames As String, qtr As String, refid As String, ByRef robj As thObject) As String
        Dim xtr As XmlTextReader = New XmlTextReader(filenames, XmlNodeType.Document, Nothing)
        xtr.WhitespaceHandling = WhitespaceHandling.None
        Dim xd As XmlDocument = New XmlDocument()
        Dim DS As New DataSet

        Dim productXML As XDocument = XDocument.Load(filenames)

        Dim productsDoc = System.Xml.Linq.XDocument.Parse(productXML.ToString())

        ' get all <Product> elements from the XDocument
        Dim products = From product In productsDoc...<Dataitems>
                       Select product
        Dim p1 = (From p In productsDoc.Elements.First.Elements.First.Elements Select p)
        Dim p2 = (From p In p1.Elements.First.Elements Select p)
        Dim p4 = (From p In p1.Elements.Last.Elements Select p)
        Dim p3 = p2.ToList
        Dim p5 = p4.ToList
        'Dim robject As New thObject
        Dim sumry As New thSummary
        With sumry
            Dim obj
            obj = Split(Split(p3.Item(2).ToString, ">")(1), "<")(0)
            .employerCode = p3.ElementAt(8).Value
            .firstname = p3.ElementAt(4).Value
            .fundCode = p3.ElementAt(50).Value
            .middlename = p3.ElementAt(5).Value
            .quarterId = qtr
            .referenceId = refid
            .rsaBalance = Math.Round(CDbl(p3.ElementAt(23).Value), 2)
            .rsapin = p3.ElementAt(3).Value
            .surname = p3.ElementAt(6).Value
            .tpfacode = "026"
            .ttlGainOrLoss = Math.Round(CDbl(p3.ElementAt(22).Value), 2)
            .ttlNoOfUnits = Math.Round(CDbl(p3.ElementAt(23).Value), 2)
            .unitPrice = Math.Round(CDbl(p3.ElementAt(28).Value), 4)
        End With
        Dim detailrecords As New List(Of detailRecord)

        For i As Integer = 0 To p4.Count - 1
            Dim trans As New detailRecord
            '
            With trans

                Dim p6 = (From p In p5.ElementAt(i).Elements Select p)

                Dim ledger = p6.ElementAt(0).Elements
                .emplContribution = Math.Round(CDbl(ledger.ElementAt(3).Value), 2)
                .emplContribution = Math.Round(CDbl(ledger.ElementAt(3).Value), 2)
                .employeeContribution = Math.Round(CDbl(ledger.ElementAt(4).Value), 2)
                .fees = Math.Round(CDbl(ledger.ElementAt(7).Value), 2)
                .netContribution = Math.Round(CDbl(ledger.ElementAt(8).Value), 2)
                .numberOfUnits = Math.Round(CDbl(ledger.ElementAt(11).Value), 4)
                .others = Math.Round(CDbl(ledger.ElementAt(14).Value.ToString), 2)
                .paymentDate = CDate(ledger.ElementAt(1).Value)
                .quarterId = qtr
                .referenceId = refid
                If ledger.ElementAt(12).Value = "" Then ledger.ElementAt(12).Value = "1753-01-01"
                If ledger.ElementAt(13).Value = "" Then ledger.ElementAt(13).Value = "1753-01-01"
                .relatedMnthEnd = CDate(ledger.ElementAt(12).Value)
                .relatedMnthStart = CDate(ledger.ElementAt(13).Value)
                .relatedPfaCode = "026"
                .serialNo = i.ToString
                .totalContribution = Math.Round(CDbl(.emplContribution + .employeeContribution + .voluntaryContingent + .voluntaryRetirement + .others), 2)
                .transactionType = ledger.ElementAt(15).Value
                .voluntaryContingent = ledger.ElementAt(5).Value
                .voluntaryRetirement = ledger.ElementAt(6).Value
                .withdrawal = ledger.ElementAt(9).Value

            End With
            detailrecords.Add(trans)

        Next

        robj.detailRecords = detailrecords
        robj.thSummary = sumry

        Return "SUCCESS"
    End Function
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

            fullerror = ex.Message
            If fullerror.ToUpper.Contains("VIOLATION") Then
                fullerror = "Price has already been imported for fundid on date"
            End If
            Return Nothing
        End Try
    End Function
    Public fullerror As String
End Class