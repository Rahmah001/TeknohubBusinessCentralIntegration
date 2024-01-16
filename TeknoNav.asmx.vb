Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Data.SqlClient

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' This webservice is for milesoft to send accounting entries and  FUND PRICE TO MICROSOFT DYNAMICS NAVISION
<System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://www.attain-es.com/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class AttainNAV
    Inherits System.Web.Services.WebService
    ' THIS METHOD IS CALLED BY MILESOFT TO PUSH FUND PRICE TO NAVISION
    <WebMethod()> _
    Public Function RequestNav(xmlstring As String) As String
        Try
            Dim newdt As New DataTable
            newdt.Columns.Add("PlanCode")
            newdt.Columns.Add("Date")
            newdt.Columns.Add("Status")
            newdt.Columns.Add("ReferenceNo")
            newdt.Columns.Add("Error_Code")
            newdt.Columns.Add("Error_Message")

            Dim ds As New DataSet
            Try
                ds.ReadXml(New StringReader(xmlstring))
            Catch ex As Exception
                Return "INVALID XML FORMAT"
            End Try

            Try
                LogReceivingSending(ds, xmlstring, "R")
            Catch ex As Exception

            End Try


            Dim dt As New DataTable
            dt = ds.Tables("DataItem")
            Dim fundid, comment As String
            Dim dDate As Date
            Dim query As String = ""
            Dim fullquery As String = ""
            Dim totresp As String = ""

            Dim bidprice, offerprice, actualnse, fundperf, marketprice, navpfa As Decimal
            For i As Integer = 0 To dt.Rows.Count - 1
                Dim dr As DataRow
                dr = newdt.NewRow
                fundid = dt.Rows(i).Item("PlanCode").ToString
                comment = ""
                'Error checking for the navdate, bid price and  offer price
                Try
                    Dim InputDateString As String = dt.Rows(i).Item("NAVDate")
                    Dim iYear As Integer = System.Convert.ToInt32(InputDateString.Substring(0, 4))
                    Dim iMonth As Integer = System.Convert.ToInt32(InputDateString.Substring(4, 2))
                    Dim iDay As Integer = System.Convert.ToInt32(InputDateString.Substring(6, 2))
                    dDate = New DateTime(iYear, iMonth, iDay)
                Catch ex As Exception
                    fullerror = "INVALID DATE DATA"

                    GoTo ERR   'For writting to the datatable on errror for the current row
                End Try

                If IsNumeric(dt.Rows(i).Item("Bidprice")) = False Or IsNumeric(dt.Rows(i).Item("OfferPrice")) = False Or IsNumeric(dt.Rows(i).Item("FundPerformance")) = False Then
                    fullerror = "INVALID NUMERIC DATA"
                    GoTo ERR  'For writting to the datatable on errror for the current row
                End If
                Try
                    bidprice = dt.Rows(i).Item("Bidprice")
                    bidprice = Math.Round(bidprice, 4)
                    If Val(bidprice) = 0 Then
                        fullerror = "INVALID BID PRICE"
                        GoTo ERR 'For writting to the datatable on errror for the current row
                    End If
                Catch ex As Exception
                    fullerror = "INVALID BID PRICE"
                    GoTo ERR 'For writting to the datatable on errror for the current row
                End Try

                Try
                    offerprice = dt.Rows(i).Item("OfferPrice")
                    offerprice = Math.Round(offerprice, 4)
                    If Val(offerprice) = 0 Then
                        fullerror = "INVALID OFFER PRICE"
                        GoTo ERR 'For writting to the datatable on errror for the current row
                    End If
                Catch ex As Exception
                    fullerror = "INVALID OFFER PRICE"
                    GoTo ERR 'For writting to the datatable on errror for the current row
                End Try
                actualnse = 0.0
                fundperf = dt.Rows(i).Item("FundPerformance")
                marketprice = 0.0
                navpfa = 0.0

                'End of validation and begining of insertion into the fund price table in navision

                Dim readquery As String = "select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Fund Price] where [Fund No_]='" & fundid & "' and [Date]= '" & dDate.ToString("yyyyMMdd") & "'"
                Dim newds As DataSet = returndataset(readquery)
                If newds Is Nothing Then
                    GoTo Err
                End If
                fullerror = ""

                If newds.Tables(0).Rows.Count > 0 Then
                    'if it exist , confirm if its the maximum date to update . otherwise raise error

                    Dim maxdateconfirm As String = "select max([Date]) from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Fund Price] where [Fund No_]='" & fundid & "'"
                    Dim confds As DataSet = returndataset(maxdateconfirm)
                    If CDate(confds.Tables(0).Rows(0).Item(0).ToString) = dDate Then
                        query = "update  [" & My.MySettings.Default.COMPANYNAME.Trim & "$Fund Price] set [Bid Price LCY]='" & bidprice & "',[Offer Price LCY]='" & offerprice & "',"
                        query = query & " [Fund Performance] = '" & fundperf & "', [Comments] ='edited on " & Today.Date & "' where [Fund No_]='" & fundid & "' and [Date]= '" & dDate.ToString("yyyyMMdd") & "'"
                    Else
                        fullerror = "Greater dates already available in the system"
                        GoTo Err 'For writting to the datatable on errror for the current row
                    End If
                Else

                    query = "Insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Fund Price] ([Fund No_],[Date],[Bid Price LCY],[Offer Price LCY],[Comments],[Actual NSE Index],"
                    query = query & "[Fund Performance],[Market Performance],[NAV (PFA)])"
                    query = query & " values('" & fundid & "','" & dDate.ToString("yyyyMMdd") & "','" & bidprice & "','" & offerprice & "','" & comment & "','" & actualnse & ""
                    query = query & "','" & fundperf & "','" & marketprice & "','" & navpfa & "')"

                End If
                returndataset(query)



                If fullerror.Trim <> "" Then
                    fullerror = Replace(fullerror, "fundid", fundid)
                    fullerror = Replace(fullerror, "date", dDate.ToString("yyyyMMdd"))
Err:                dr.Item("PlanCode") = fundid
                    Try
                        dr.Item("Date") = dDate.ToString("yyyyMMdd")
                    Catch ex As Exception

                    End Try

                    dr.Item("Status") = "F"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = geterrorcode(fullerror)
                    dr.Item("Error_Message") = fullerror
                    newdt.Rows.Add(dr)
                Else
                    dr.Item("PlanCode") = fundid
                    dr.Item("Date") = dDate.ToString("yyyyMMdd")
                    dr.Item("Status") = "S"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = ""
                    dr.Item("Error_Message") = ""
                    newdt.Rows.Add(dr)
                End If
            Next
            Dim dsres As New DataSet
            newdt.TableName = "DataItem"
            dsres.DataSetName = "Reply"
            dsres.Tables.Add(ds.Tables("Header").Copy)
            dsres.Tables.Add(newdt)

            Try
                LogReceivingSending(dsres, ToStringAsXml(dsres), "S")
            Catch ex As Exception

            End Try

            Return ToStringAsXml(dsres)
        Catch ex As Exception

            Return ex.Message
        End Try
    End Function
    'function to convert date to nav date format
    Private Function getdbdate(inputdatestring As String) As Date
        If IsDate(inputdatestring) Then
            Return CDate(inputdatestring)
            Exit Function
        End If
        If inputdatestring = "" Then inputdatestring = "17530101"
        Dim iYear As Integer = System.Convert.ToInt32(inputdatestring.Substring(0, 4))
        Dim iMonth As Integer = System.Convert.ToInt32(inputdatestring.Substring(4, 2))
        Dim iDay As Integer = System.Convert.ToInt32(inputdatestring.Substring(6, 2))
        getdbdate = New DateTime(iYear, iMonth, iDay)
        Return getdbdate
    End Function

    Dim fullerror As String
    ' for executing ms sql server query 
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
    'functions to get equivalent error code
    Private Function geterrorcode(ByVal errmesg As String) As String
        geterrorcode = errmesg.ToUpper.Trim

        If errmesg.ToUpper.Contains("ALREADY") Then
            Return "NAVER0001"
        End If
        If errmesg.ToUpper.Contains("NUMERIC") Then
            Return "NAVER0002"
        End If
        If errmesg.ToUpper.Contains("DATE") Then
            Return "NAVER0003"
        End If
        If errmesg.ToUpper = "INVALID OFFER PRICE" Then
            Return "NAVER0004"
        End If
        If errmesg.ToUpper = "INVALID BID PRICE" Then
            Return "NAVER0005"
        End If
        Return "NAVER0010"
    End Function
    Dim oldvouchnum As New ArrayList
    <WebMethod()>
    Public Function RequestAccountingEntries(xmlstring As String) As String
        ' THIS METHOD IS CALLED BY MILESOFT TO PUSH ACCOUNTING ENTRIES  TO COMPANY ACCOUNT IN NAVISION
        Try
            Dim newdt As New DataTable
            newdt.Columns.Add("VoucherNumber")
            newdt.Columns.Add("LedgerAccountCode")
            newdt.Columns.Add("Status")
            newdt.Columns.Add("ReferenceNo")
            newdt.Columns.Add("Error_Code")
            newdt.Columns.Add("Error_Message")

            Dim ds As New DataSet
            Try
                ds.ReadXml(New StringReader(xmlstring))
            Catch ex As Exception
                Return "INVALID XML FORMAT"
            End Try



            Dim txnrefnum As String = ds.Tables(0).Rows(0).Item("TxnRefNum").ToString
            Dim txntime As String = ds.Tables(0).Rows(0).Item("TxnTime").ToString
            Dim txndate As Date

            Dim InputDateString As String
            Try
                InputDateString = ds.Tables(0).Rows(0).Item("TxnDate").ToString
                Dim iYear As Integer = System.Convert.ToInt32(InputDateString.Substring(0, 4))
                Dim iMonth As Integer = System.Convert.ToInt32(InputDateString.Substring(4, 2))
                Dim iDay As Integer = System.Convert.ToInt32(InputDateString.Substring(6, 2))
                txndate = New DateTime(iYear, iMonth, iDay)
            Catch ex As Exception
                ' fullerror = "INVALID DATE DATA"
                Return "INVALID DATE DATA FOR TXNDATE"
                ' GoTo Err
            End Try
            savetoFundawareResponses(txnrefnum, txntime, InputDateString, xmlstring, "R")

            Dim dt As New DataTable
            Try
                dt = ds.Tables("DataItem")
            Catch ex As Exception
                Return " There are no Data item in your xml. Kindly review and resend"
            End Try

            Dim comment As String
            Dim query As String = ""
            Dim fullquery As String = ""
            Dim totresp As String = ""
            Dim fundcode As String

            Dim accountingdate As Date
            Dim ledgeraccountCODE As String
            Dim dr As DataRow
            If dt Is Nothing Then
                Return " There are no Data item in your xml. Kindly review and resend"
            End If
            If dt.Rows.Count = 0 Then
                Return " There are no Data item in your xml. Kindly review and resend"
            End If
            'If isequally(dt) = False Then
            '    fullerroracc = "One or More of the voucher number's debit and credit does not match"
            '    For i As Integer = 0 To dt.Rows.Count - 1
            '        dr = newdt.NewRow
            '        dr.Item("VoucherNumber") = dt.Rows(i).Item("VoucherNumber")
            '        dr.Item("LedgerAccountCode") = dt.Rows(i).Item("LedgerAccountCode").ToString
            '        dr.Item("Status") = "F"
            '        dr.Item("ReferenceNo") = ""
            '        dr.Item("Error_Code") = ""
            '        dr.Item("Error_Message") = fullerroracc
            '        newdt.Rows.Add(dr)
            '    Next
            '    GoTo prosend
            'End If
            If isVerified(dt) = False Then
                fullerroracc = "One or More of the Account code does not exists (e.g." & getacccode & ")"
                For i As Integer = 0 To dt.Rows.Count - 1
                    dr = newdt.NewRow
                    dr.Item("VoucherNumber") = dt.Rows(i).Item("VoucherNumber")
                    dr.Item("LedgerAccountCode") = dt.Rows(i).Item("LedgerAccountCode").ToString
                    dr.Item("Status") = "F"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = ""
                    dr.Item("Error_Message") = fullerroracc
                    newdt.Rows.Add(dr)
                Next
                GoTo prosend
            End If






            If insertRecievedRecordstoDB(ds) = False Then
                'fullerroracc = "One or More of the Account code does not exists (e.g." & getacccode & ")"
                For i As Integer = 0 To dt.Rows.Count - 1
                    dr = newdt.NewRow
                    dr.Item("VoucherNumber") = dt.Rows(i).Item("VoucherNumber")
                    dr.Item("LedgerAccountCode") = dt.Rows(i).Item("LedgerAccountCode").ToString
                    dr.Item("Status") = "F"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = ""
                    dr.Item("Error_Message") = fullerroracc
                    newdt.Rows.Add(dr)
                Next
                GoTo prosend
            End If





            For i As Integer = 0 To dt.Rows.Count - 1
                If fullerroracc <> "" Then
                    dr = newdt.NewRow
                    fullerroracc = " Voucher Number previously imported to the system or  some record(s)  awaits posting."
                    GoTo Err
                End If

                fullerroracc = ""
                dr = newdt.NewRow
                fundcode = dt.Rows(i).Item("PlanCode").ToString
                comment = ""
                Try
                    InputDateString = dt.Rows(i).Item("AccountingDate")
                    Dim iYear As Integer = System.Convert.ToInt32(InputDateString.Substring(0, 4))
                    Dim iMonth As Integer = System.Convert.ToInt32(InputDateString.Substring(4, 2))
                    Dim iDay As Integer = System.Convert.ToInt32(InputDateString.Substring(6, 2))
                    accountingdate = New DateTime(iYear, iMonth, iDay)
                Catch ex As Exception
                    fullerroracc = "INVALID DATE DATA"
                    GoTo Err
                End Try

                If IsNumeric(dt.Rows(i).Item("Amount")) = False Then
                    fullerroracc = "INVALID NUMERIC DATA"
                    GoTo Err
                End If
                Try
                    If Val(dt.Rows(i).Item("Amount")) <= 0 Then
                        fullerroracc = "INVALID AMOUNT"
                        GoTo Err
                    End If
                Catch ex As Exception
                    fullerroracc = "INVALID AMOUNT"
                    GoTo Err
                End Try

                Try
                    ledgeraccountCODE = dt.Rows(i).Item("LedgerAccountCode")
                    If ledgeraccountCODE.Trim.Length = 0 Then
                        fullerroracc = "INVALID LEDGER ACCOUNT CODE"
                        GoTo Err
                    End If
                Catch ex As Exception
                    fullerroracc = "INVALID LEDGER ACCOUNT CODE"
                    GoTo Err
                End Try
                fullerroracc = ""




                If fullerroracc.Trim <> "" Then
                    'IF THERE IS ERROR  IT COMES TO THIS POINT 
Err:                dr.Item("VoucherNumber") = dt.Rows(i).Item("VoucherNumber").ToString
                    dr.Item("LedgerAccountCode") = dt.Rows(i).Item("LedgerAccountCode").ToString
                    dr.Item("Status") = "F"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = ""
                    dr.Item("Error_Message") = fullerroracc
                    newdt.Rows.Add(dr)
                Else
                    dr.Item("VoucherNumber") = dt.Rows(i).Item("VoucherNumber").ToString
                    dr.Item("LedgerAccountCode") = dt.Rows(i).Item("LedgerAccountCode").ToString
                    dr.Item("Status") = "S"
                    dr.Item("ReferenceNo") = ""
                    dr.Item("Error_Code") = ""
                    dr.Item("Error_Message") = ""
                    newdt.Rows.Add(dr)
                End If

                '  savetoglAccountTable(deriveddt.Rows(i))
            Next

            If fullerroracc.Trim = "" Then
                Try
                    Dim indt As New DataTable
                    indt = getTotalperledgercode(deriveddt)
                    For i As Integer = 0 To dt.Rows.Count - 1
                        manipulate_saveforGeneric(deriveddt.Rows(i))
                    Next
                    createGenJournalBatch(journalBatchName)
                Catch ex As Exception
                    WriteLine(ex.Message)
                End Try
            End If


prosend:    Dim dsres As New DataSet
            newdt.TableName = "DataItem"
            dsres.DataSetName = "Reply"
            dsres.Tables.Add(ds.Tables("Header").Copy)
            dsres.Tables.Add(newdt)
            Dim sentresp As String = ToStringAsXml(dsres)
            savetoFundawareResponses(txnrefnum, txntime, InputDateString, sentresp, "S")



            Return sentresp
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
    Public constantcount As Integer
    Dim distinctDT As DataTable
    Public journalBatchName As String
    Private Function isVerified(ByVal mydt As DataTable) As Boolean
        'FUNCTION TO CONFIRM THE EXISTENCE OF THE ACCOUNT CODE FROM GL ACCOUNT TABLE IN NAV
        Try
            Dim ndt As New DataTable
            ndt.Columns.Add("LedgerAccountCode")
            ndt.Merge(mydt)

            distinctDT = ndt.DefaultView.ToTable(True, "LedgerAccountCode")

            Dim qstring As String = ""
            For i As Integer = 0 To distinctDT.Rows.Count - 1
                Dim db As New accountingdDataContext
                qstring = distinctDT.Rows(i).Item("LedgerAccountCode").ToString
                'Dim rec = (From k In db.ARM_COMPANY_ACCOUNTS_G_L_Accounts Where k.No_ = qstring).FirstOrDefault
                'If rec Is Nothing Then
                '    getacccode = qstring
                '    Return False
                'End If
                Dim ds As New DataSet
                getacccode = qstring
                ds = returnfulldataset("Select * from [ARM COMPANY ACCOUNTS$Fundware FEES Mapping Codes] where code ='" & qstring & "'")
                If ds Is Nothing Then Return False
                If ds.Tables.Count = 0 Then Return False
                If ds.Tables(0).Rows.Count = 0 Then Return False

            Next
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function
    Private Function insertRecievedRecordstoDB(ByVal ds As DataSet) As Boolean
        oldvouchnum = New ArrayList
        insertRecievedRecordstoDB = False
        Dim con As New SqlConnection
        con.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
        con.Open()
        Dim dt As DataTable = ds.Tables("dataitem")




        Try
            ' Write from the source to the destination.
            Dim txnref As String = ds.Tables(0).Rows(0).Item("TxnRefNum").ToString


            dt.Columns.Add("Status", Type.GetType("System.Int32"), 0)
            Dim dc As DataColumn
            dc = New DataColumn("TxnDate", Type.GetType("System.DateTime"))
            dc.DefaultValue = getdbdate(ds.Tables(0).Rows(0).Item("TxnDate"))
            dt.Columns.Add(dc)
            dc = New DataColumn("TxnTime", Type.GetType("System.String"))
            dc.DefaultValue = ds.Tables(0).Rows(0).Item("TxnTime")
            dt.Columns.Add(dc)
            dc = New DataColumn("TxnRefNum", Type.GetType("System.String"))
            dc.DefaultValue = txnref
            dt.Columns.Add(dc)
            dt.Columns.Add("Counter", Type.GetType("System.Int32"))

            For i As Integer = 0 To dt.Rows.Count - 1
                dt.Rows(i).Item("Counter") = i + 1
            Next
            deriveddt = dt
            Using copy As New SqlBulkCopy(con)
                copy.ColumnMappings.Add("TxnRefNum", "TxnRefNum")
                copy.ColumnMappings.Add("[SchemeCode]", "SchemeCode")
                copy.ColumnMappings.Add("TxnTime", "TxnTime")
                copy.ColumnMappings.Add("OptionCode", "OptionCode")
                copy.ColumnMappings.Add("LedgerAccountCode", "LedgerAccountCode")
                copy.ColumnMappings.Add("VoucherNumber", "VoucherNumber")
                copy.ColumnMappings.Add("AccountingDate", "AccountingDate")
                copy.ColumnMappings.Add("PlanCode", "PlanCode")
                copy.ColumnMappings.Add("Amount", "Amount")
                copy.ColumnMappings.Add("CurrencyShortName", "CurrencyShortName")
                copy.ColumnMappings.Add("DebitCredit", "DebitCredit")
                copy.ColumnMappings.Add("FXRate", "FXRate")
                copy.ColumnMappings.Add("SecurityISIN", "SecurityISIN")
                copy.ColumnMappings.Add("Narration", "Narration")
                copy.ColumnMappings.Add("TxnDate", "TxnDate")
                copy.ColumnMappings.Add("Status", "Status")
                copy.ColumnMappings.Add("Counter", "Counter")
                copy.DestinationTableName = "[ARM COMPANY ACCOUNTS$Fundware Received Account Enty]"
                copy.WriteToServer(dt)
            End Using
            insertRecievedRecordstoDB = True
        Catch ex As Exception
            insertRecievedRecordstoDB = False
            If ex.Message.ToString.ToLower.Contains("violation") Then
                insertRecievedRecordstoDB = True
            End If
        End Try



        Return insertRecievedRecordstoDB




    End Function
    Public deriveddt As DataTable
    Public getacccode As String
    Private Function isequally(ByVal mydt As DataTable) As Boolean
        'FUNCTION TO CONFIRM IF THE CREDIT = DEBIT FOR EACH VOUCHER NUMBER
        Try
            Dim distinctDT As DataTable = mydt.DefaultView.ToTable(True, "VoucherNumber")
            Dim dv As New DataView
            Dim mysumDebit, mysumCredit As Decimal
            dv = mydt.DefaultView
            Dim qstring As String = ""
            For i As Integer = 0 To distinctDT.Rows.Count - 1
                qstring = "VoucherNumber ='" & distinctDT.Rows(i).Item("VoucherNumber") & "' AND DebitCredit ='DEBIT'"
                dv.RowFilter = (qstring)

                For k As Integer = 0 To dv.ToTable.Rows.Count - 1
                    mysumDebit = mysumDebit + CDbl(dv.ToTable.Rows(k).Item("Amount"))
                Next

                qstring = "VoucherNumber ='" & distinctDT.Rows(i).Item("VoucherNumber") & "' AND DebitCredit ='CREDIT'"
                dv.RowFilter = (qstring)

                For k As Integer = 0 To dv.ToTable.Rows.Count - 1
                    mysumCredit = mysumCredit + CDbl(dv.ToTable.Rows(k).Item("Amount"))
                Next

                If mysumCredit <> mysumDebit Then
                    Return False
                End If
            Next
        Catch ex As Exception
            Return False
        End Try
        Return True
    End Function

    Private Function getTotalperledgercode(ByVal dt As DataTable) As DataTable
        Dim newdt As New DataTable
        newdt.Columns.Add("LedgerAccountCode")
        newdt.Columns.Add("Amount")
        newdt.Columns.Add("VoucherNumber")
        'FUNCTION TO CONFIRM IF THE CREDIT = DEBIT FOR EACH VOUCHER NUMBER
        Try
            Dim names = From row In dt.AsEnumerable()
                        Select row.Field(Of String)("LedgerAccountCode") Distinct


            '  Dim distinctDT As DataTable = dt.DefaultView.ToTable(False, "LedgerAccountCode")
            Dim dv As New DataView
            Dim mysum As Decimal
            Dim vouchnum As String = ""
            dv = dt.DefaultView
            Dim qstring As String = ""
            For i As Integer = 0 To names.Count - 1
                qstring = "LedgerAccountCode ='" & names(i).ToString & "' "
                dv.RowFilter = (qstring)

                For k As Integer = 0 To dv.ToTable.Rows.Count - 1
                    mysum = mysum + CDbl(dv.ToTable.Rows(k).Item("Amount"))
                Next

                vouchnum = dv.Table.Rows(i).Item("VoucherNumber").ToString

                Dim dr As DataRow
                dr = newdt.NewRow
                dr.Item(0) = names(i).ToString
                dr.Item(1) = mysum
                dr.Item(2) = vouchnum
                newdt.Rows.Add(dr)

            Next
        Catch ex As Exception
            Return Nothing
        End Try
        Return (newdt)

    End Function
    Private Function manipulate_saveforConstant(ByVal dr As DataRow) As Boolean
        Dim ds As New DataSet
        ds = returnfulldataset("Select * from  [ARM COMPANY ACCOUNTS$GL Fundware VAT-WHT Codes] where [Active] = 1")
        Dim k As Integer = constantcount



        If ds Is Nothing Then Return False
        If ds.Tables.Count = 0 Then Return False
        If ds.Tables(0).Rows.Count = 0 Then Return True
        constantcount = ds.Tables(0).Rows.Count - 1
        For j As Integer = 0 To ds.Tables(0).Rows.Count - 1
            Dim myamt = (CDbl(dr.Item("Amount")) / 105) * CDbl(ds.Tables(0).Rows(j).Item("Percentage"))
            savetoglAccountTable(ds.Tables(0).Rows(j).Item("code").ToString, ds.Tables(0).Rows(j).Item("Debit_Credit").ToString, myamt, getdbdate(deriveddt.Rows(0).Item("TxnDate")),
                                 dr.Item("TxnRefNum").ToString, k, getdbdate(dr.Item("AccountingDate")), dr.Item("Narration").ToString, dr.Item("VoucherNumber").ToString)
            k += 1
        Next

        constantcount = k
        Return True
    End Function
    Private Function manipulate_saveforGeneric(ByVal dr As DataRow) As Boolean
        Dim ds As New DataSet
        '------------just edited on 10 nov. 2015'
        If dr.Item("DebitCredit").ToString.ToUpper.StartsWith("C") Then
            dr.Item("Amount") = -1 * CDbl(dr.Item("Amount"))
        End If
        '------------just edited on 10 nov. 2015 END'


        Dim k As Integer = constantcount
        '   For i As Integer = 0 To dt.Rows.Count - 1
        ds = returnfulldataset("Select * from [ARM COMPANY ACCOUNTS$GL Fundware Mapping Codes] where [GL Attached] ='" & dr.Item("LedgerAccountCode") & "' and [Active] = 1")
        If ds Is Nothing = False Then 'Then Return False
            If ds.Tables.Count = 0 Then Return False
            If ds.Tables(0).Rows.Count = 0 Then Return False
            For j As Integer = 0 To ds.Tables(0).Rows.Count - 1
                Dim myamt = (CDbl(dr.Item("Amount")) / 105) * CDbl(ds.Tables(0).Rows(j).Item("Percentage"))
                savetoglAccountTable(ds.Tables(0).Rows(j).Item("code").ToString, ds.Tables(0).Rows(j).Item("Debit_Credit").ToString, myamt, getdbdate(dr.Item("TxnDate")),
                                     dr.Item("TxnRefNum").ToString, k, getdbdate(dr.Item("AccountingDate")), dr.Item("Narration").ToString, dr.Item("VoucherNumber").ToString)
                k += 1
            Next
            constantcount = k
            manipulate_saveforConstant(dr)
        End If
        ' Next

        Return True
    End Function
    Private Function createGenJournalBatch(ByVal batchName As String) As Boolean


        Try


            ' TO CREATE A NEW BATCH TO THE GEN. JOURNAL BATCH TABLE

            Dim db As New accountingdDataContext
            Dim rec As New ARM_COMPANY_ACCOUNTS_Gen__Journal_Batch

            rec.Allow_VAT_Difference = 0
            rec.Allow_Payment_Export = 0

            rec.Bal__Account_No_ = ""
            rec.Bank_Statement_Import_Format = ""
            rec.Suggest_Balancing_Amount = 0

            rec.Bal__Account_Type = 0
            rec.Copy_VAT_Setup_to_Jnl__Lines = 1
            rec.Description = batchName
            rec.Journal_Template_Name = "GENERAL"
            rec.Name = batchName
            rec.No__Series = ""
            rec.Posting_No__Series = ""
            rec.Reason_Code = ""

            db.ARM_COMPANY_ACCOUNTS_Gen__Journal_Batches.InsertOnSubmit(rec)
            db.SubmitChanges()
            Return True
        Catch ex As Exception
            fullerroracc = ex.Message
            Return False
        End Try
        Return True
    End Function
    Private Function savetoglAccountTable(accCode As String, dc As String, amt As Double, txndate As Date, txnrefno As String, counter As Int32,
                                         accDate As Date, narration As String, vouchnum As String) As Boolean

        journalBatchName = txnrefno.Trim
        amt = Math.Round(amt, 2)
        Try
            ' TO SAVE THE RECEIVED ACCOUNTING ENTRIES TO GEN JOURNAL LINE TABLE 

            Dim db As New accountingdDataContext
            Dim MYREC As New ARM_COMPANY_ACCOUNTS_Gen__Journal_Line
            MYREC.Account_Name = ""
            MYREC.Account_No_ = accCode
            MYREC.Account_Type = 0
            MYREC.Additional_Currency_Posting = 0
            MYREC.Agent_Code = ""
            MYREC.Allow_Application = 0
            MYREC.Allow_Zero_Amount_Posting = 0
            MYREC.Amount__LCY_ = 0
            MYREC.Applies_to_Doc__No_ = ""
            MYREC.Applies_to_Doc__Type = 0
            MYREC.Applies_to_ID = ""
            MYREC.Bal__Account_Name = ""
            MYREC.Bal__Account_No_ = ""
            MYREC.Bal__Account_Type = 0
            MYREC.Bal__Gen__Bus__Posting_Group = ""
            MYREC.Bal__Gen__Posting_Type = 0
            MYREC.Bal__Gen__Prod__Posting_Group = ""
            MYREC.Bal__Tax_Area_Code = ""
            MYREC.Bal__Tax_Group_Code = ""
            MYREC.Bal__Tax_Liable = 0
            MYREC.Bal__Use_Tax = 0
            MYREC.Bal__VAT__ = 0
            MYREC.Bal__VAT_Amount__LCY_ = 0
            MYREC.Bal__VAT_Amount = 0
            MYREC.Bal__VAT_Base_Amount__LCY_ = 0
            MYREC.Bal__VAT_Base_Amount = 0
            MYREC.Bal__VAT_Bus__Posting_Group = ""
            MYREC.Bal__VAT_Calculation_Type = 0
            MYREC.Bal__VAT_Difference = 0
            MYREC.Bal__VAT_Prod__Posting_Group = ""
            MYREC.Balance__LCY_ = 0
            MYREC.Bank_Entry_No = 0
            MYREC.Bank_Payment_Type = 0
            MYREC.Bank_Statement_Line_No = 0
            MYREC.Basic_Salary = 0
            MYREC.Bill_to_Pay_to_No_ = ""
            MYREC.Budgeted_FA_No_ = ""
            MYREC.Business_Unit_Code = ""
            MYREC.Campaign_No_ = ""
            MYREC.Check_Printed = 0
            MYREC.Cheque_Date = getdbdate("")
            MYREC.Contribution_Source = ""
            MYREC.Contribution_Type = 0
            MYREC.Contribution_Types = ""
            MYREC.Contrubtion_Period = getdbdate("")
            MYREC.Country_Region_Code = ""
            'If dc.ToString.ToUpper.StartsWith("C") Then
            '    MYREC.Credit_Amount = amt
            '    MYREC.Debit_Amount = 0
            '    MYREC.Amount = (-1) * amt
            'Else
            '    MYREC.Credit_Amount = 0
            '    MYREC.Debit_Amount = amt
            '    MYREC.Amount = amt
            'End If
            If dc = 1 Then 'credit
                MYREC.Credit_Amount = amt
                MYREC.Debit_Amount = 0
                MYREC.Amount = (-1) * amt
            End If

            If dc = 0 Then 'debit
                MYREC.Credit_Amount = 0
                MYREC.Debit_Amount = amt
                MYREC.Amount = amt
            End If
            MYREC.Credit_Card_No_ = ""
            MYREC.Currency_Code = ""
            MYREC.Currency_Factor = 0
            MYREC.Depr__Acquisition_Cost = 0
            MYREC.Depr__until_FA_Posting_Date = 0
            MYREC.Depreciation_Book_Code = ""
            MYREC.Dimension_Set_ID = 0
            MYREC.Document_Date = txndate
            MYREC.Document_No_ = txnrefno
            MYREC.Document_Type = 0
            MYREC.Due_Date = getdbdate("")
            MYREC.Duplicate_in_Depreciation_Book = ""
            MYREC.Employer_No = ""
            MYREC.EU_3_Party_Trade = 0
            MYREC.Exemption_Type = 0
            MYREC.Expiration_Date = getdbdate("")
            MYREC.External_Document_No_ = ""
            MYREC.FA_Add__Currency_Factor = 0
            MYREC.FA_Error_Entry_No_ = 0
            MYREC.FA_Posting_Date = getdbdate("")
            MYREC.FA_Posting_Type = 0
            MYREC.FA_Reclassification_Entry = 0
            MYREC.Financial_Void = 0
            MYREC.Fund_Code = ""
            MYREC.Gen__Bus__Posting_Group = ""
            MYREC.Gen__Posting_Type = 0
            MYREC.Gen__Prod__Posting_Group = ""
            MYREC.IC_Direction = 0
            MYREC.IC_Partner_Code = ""
            MYREC.IC_Partner_G_L_Acc__No_ = ""
            MYREC.IC_Partner_Transaction_No_ = 0
            MYREC.Index_Entry = 0
            MYREC.Insurance_No_ = ""
            MYREC.Inv__Discount__LCY_ = 0
            MYREC.Job_Currency_Code = ""
            MYREC.Job_Currency_Factor = 0
            MYREC.Job_Line_Amount__LCY_ = 0
            MYREC.Job_Line_Amount = 0
            MYREC.Job_Line_Disc__Amount__LCY_ = 0
            MYREC.Job_Line_Discount__ = 0
            MYREC.Job_Line_Discount_Amount = 0
            MYREC.Job_Line_Type = 0
            MYREC.Job_No_ = ""
            MYREC.Job_Planning_Line_No_ = 0
            MYREC.Job_Quantity = 0
            MYREC.Job_Remaining_Qty_ = 0
            MYREC.Job_Task_No_ = ""
            MYREC.Job_Total_Cost__LCY_ = 0
            MYREC.Job_Total_Cost = 0
            MYREC.Job_Total_Price__LCY_ = 0
            MYREC.Job_Total_Price = 0
            MYREC.Job_Unit_Cost__LCY_ = 0
            MYREC.Job_Unit_Cost = 0
            MYREC.Job_Unit_Of_Measure_Code = ""
            MYREC.Job_Unit_Price__LCY_ = 0
            MYREC.Job_Unit_Price = 0
            MYREC.Journal_Batch_Name = journalBatchName
            MYREC.Journal_Template_Name = "GENERAL"
            MYREC.Line_No_ = (counter + 1) * 10000
            MYREC.Loan_No = 0
            MYREC.Maintenance_Code = ""
            MYREC.No__of_Depreciation_Days = 0
            MYREC.No__Of_Units = 0
            MYREC.On_Hold = ""
            MYREC.Payment_Discount__ = 0
            MYREC.Payment_Mode = ""
            MYREC.Payment_Terms_Code = ""
            MYREC.Period_Narration = ""
            MYREC.Pmt__Discount_Date = getdbdate("")
            MYREC.Posting_Date = accDate
            MYREC.Posting_Group = ""
            MYREC.Posting_No__Series = ""
            MYREC.Price_Per_Unit = 0
            MYREC.Prod__Order_No_ = ""
            MYREC.Profit__LCY_ = 0
            MYREC.Reason_Code = ""
            MYREC.Recurring_Frequency = ""
            MYREC.Recurring_Method = 0
            MYREC.Region_Code = ""
            MYREC.Reversing_Entry = 0
            MYREC.Sales_Purch___LCY_ = 0
            MYREC.Salespers__Purch__Code = ""
            MYREC.Salvage_Value = 0
            MYREC.Sell_to_Buy_from_No_ = ""
            MYREC.Ship_to_Order_Address_Code = ""
            MYREC.Shortcut_Dimension_1_Code = ""
            MYREC.Shortcut_Dimension_2_Code = ""
            MYREC.Source_Code = ""
            MYREC.Source_Curr__VAT_Amount = 0
            MYREC.Source_Curr__VAT_Base_Amount = 0
            MYREC.Source_Currency_Amount = 0
            MYREC.Source_Currency_Code = ""
            MYREC.Source_Line_No_ = 0
            MYREC.Source_No_ = ""
            MYREC.Source_Type = 0
            MYREC.State_Code = ""
            MYREC.Subscription_Ref = ""
            MYREC.System_Created_Entry = 0
            MYREC.Tax_Area_Code = ""
            MYREC.Tax_Group_Code = ""
            MYREC.Tax_Liable = 0
            MYREC.Transaction_Type = 0
            MYREC.Use_Duplication_List = 0
            MYREC.Use_Tax = 0
            MYREC.VAT__ = 0
            MYREC.VAT_Amount__LCY_ = 0
            MYREC.VAT_Amount = 0
            MYREC.VAT_Base_Amount__LCY_ = 0
            MYREC.VAT_Base_Amount = 0
            MYREC.VAT_Base_Discount__ = 0
            MYREC.VAT_Bus__Posting_Group = ""
            MYREC.VAT_Calculation_Type = 0
            MYREC.VAT_Difference = 0
            MYREC.VAT_Posting = 0
            MYREC.VAT_Prod__Posting_Group = ""
            MYREC.VAT_Registration_No_ = ""

            MYREC.Correction = 0
            MYREC.DC = ""
            MYREC.Description = Mid(narration & "  " & vouchnum, 1, 99)
            MYREC.Interest = 0
            MYREC.Prepayment = 0
            MYREC.Quantity = 0
            MYREC.Shortcut_Dimension_3_Code = ""
            MYREC.Shortcut_Dimension_4_Code = ""
            MYREC.Bank_ID = ""
            MYREC.Incoming_Document_Entry_No_ = 0

            MYREC.Creditor_No_ = ""
            MYREC.Payment_Reference = ""
            MYREC.Payment_Method_Code = ""
            MYREC.Applies_to_Doc__No_ = ""
            MYREC.Recipient_Bank_Account = ""
            MYREC.Message_to_Recipient = ""
            MYREC.Exported_to_Payment_File = 0
            MYREC.Applies_to_Ext__Doc__No_ = ""
            MYREC.Direct_Debit_Mandate_ID = ""
            MYREC.Payer_Information = ""
            MYREC.Deferral_Code = ""
            MYREC.Deferral_Line_No_ = 0
            MYREC.Comment = ""
            MYREC.Transaction_Information = ""
            db.ARM_COMPANY_ACCOUNTS_Gen__Journal_Lines.InsertOnSubmit(MYREC)

            db.SubmitChanges()
        Catch ex As Exception
            If ex.Message.Contains("Violation") Then
                fullerroracc = " Voucher Number previously imported to the system or  some record(s)  awaits posting."

            Else
                fullerroracc = ex.Message
            End If
            Return False
        End Try
        Return True

    End Function

    Private Function savetoFundawareResponses(ByVal txtrefnum As String, ByVal txntime As String, ByVal txndate As String, ByVal respxml As String, ByVal sr As String) As String
        'FUNCTION TO SAVE XML STRING FOR REFERENCE PURPOSES ONLY
        savetoFundawareResponses = ""
        Try
            Dim dDate As Date
            Dim InputDateString As String = txndate
            Dim iYear As Integer = System.Convert.ToInt32(InputDateString.Substring(0, 4))
            Dim iMonth As Integer = System.Convert.ToInt32(InputDateString.Substring(4, 2))
            Dim iDay As Integer = System.Convert.ToInt32(InputDateString.Substring(6, 2))
            dDate = New DateTime(iYear, iMonth, iDay)

            Dim db As New accountingdDataContext
            Dim funacc As New ARM_COMPANY_ACCOUNTS_Fundware_Accounting_xml
            funacc.TxnDate = dDate
            funacc.TxnTime = txntime
            funacc.TxnRefNum = txtrefnum
            funacc.Sent_Received = sr


            ' Dim abdata As Byte()
            '  abdata = System.Text.Encoding.Default.GetBytes(respxml)
            funacc.XmlString = respxml
            db.ARM_COMPANY_ACCOUNTS_Fundware_Accounting_xmls.InsertOnSubmit(funacc)
            db.SubmitChanges()
        Catch ex As Exception
            Return ex.Message
        End Try
        Return savetoFundawareResponses
    End Function
    Public Shared Function ToStringAsXml(ds As DataSet) As String
        'FUNCTION TO CONVERT DATASET TO XML STRING 
        Dim sw As New StringWriter()
        ds.WriteXml(sw, XmlWriteMode.IgnoreSchema)
        Dim s As String = sw.ToString()
        s = s.Replace("</Header>", "</Header>" & vbCrLf & "<DataBlock>" & vbCrLf & "<Data>")
        s = s.Replace("</Reply>", "</Data> " & vbCrLf & "</DataBlock>" & vbCrLf & "</Reply>")
        Return s
    End Function
    Dim fullerroracc As String
    Private Function returnfulldataset(ByVal query As String) As DataSet
        'FUNCTION TO RUN SQL QUERY
        returnfulldataset = Nothing
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

            fullerroracc = ex.Message
            If fullerroracc.ToUpper.Contains("VIOLATION") Then
                fullerroracc = "Price has already been imported for fundid on date"
            End If
            Return Nothing
        End Try
    End Function

    Private Sub LogReceivingSending(ds As DataSet, XMLvalule As String, stat As String)
        Dim query As String
        query = "insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$FundwareNAV  Responses] (TxnRefNum,TxnDate,TxnTime,XmlString,Status,Error_Code,Sent_Received)" &
            " Values( '" & ds.Tables(0).Rows(0).Item("TxnRefNum").ToString & "', '" & Today.ToString("yyyyMMdd") & "', '" & Today.ToShortTimeString & "', '" & XMLvalule & "','','','" & stat & "')"

        Dim nds As DataSet = returndataset(query)

    End Sub
    <WebMethod()> _
    Public Function RepushDepostiedRecord(voucherNumber As String) As String
        ' THIS METHOD IS CALLED to repush logged accounting entries to gen journal line table by rahmot on 07-04-2016
      


         
            Dim dt As New DataTable
          

            Dim query As String = ""
            Dim fullquery As String = ""
            Dim totresp As String = ""

    

        Dim ds As New DataSet

        query = "select * from [ARM COMPANY ACCOUNTS$Fundware Received Account Enty] where VoucherNumber ='" & voucherNumber & "'"

        ds = returndataset(query)
        If ds.Tables.Count = 0 Then Return "No record for this voucher in the Log table"
        If ds.Tables(0).Rows.Count = 0 Then Return "No record(s) for this voucher in the log table"

        dt = ds.Tables(0)
        '  getderiveddt(dt, dt.Rows(0).Item("TxnRefNum").ToString, dt.Rows(0).Item("TxnDate").ToString)
        Try


            Try
                Dim indt As New DataTable
                indt = getTotalperledgercode(dt)
                For i As Integer = 0 To dt.Rows.Count - 1
                    manipulate_saveforGeneric(dt.Rows(i))
                Next
                createGenJournalBatch(journalBatchName)
            Catch ex As Exception
                Return "Error: " & ex.Message
                ' WriteLine(ex.Message)
            End Try
           




            Return ""
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function
  
End Class
