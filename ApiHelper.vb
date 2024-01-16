Imports Newtonsoft.Json
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports AiicoAssistance.QuickType
Imports RestSharp
Imports RestSharp.Authenticators
Imports System.Net
Imports System.IO
Imports System.Xml.Serialization
Imports System.Data.SqlClient
Imports System.Reflection

Namespace RestSharp.Helper
    Public Class ApiHelper
        Public Function SubmitBalances(baseUrl As String, path As String, recordheader As ProvBalheader, obj As ProvObject) As String
            'Dim record = New Record()
            Try


                Dim client As RestClient = New RestClient(baseUrl + path)
                Dim jsonheader = JsonConvert.SerializeObject(recordheader)
                Dim json = JsonConvert.SerializeObject(obj)
                Dim request = New RestRequest("", Method.POST)
                request.AddHeader("token", recordheader.token)
                request.AddHeader("userid", recordheader.userid)
                request.AddHeader("password", recordheader.password)
                request.AddHeader("quarter", recordheader.quarter)

                request.AddParameter("application/json; charset=utf-8", json, ParameterType.RequestBody)
                request.RequestFormat = DataFormat.Json


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
                Dim Response As IRestResponse = client.Execute(request)
                '   Return "AFTER EXECUTION STAGE" & Response.IsSuccessful & Response.ErrorMessage & "-----" & json
                If Response.IsSuccessful = True Then
                    Dim strMyJson1 As String = Response.Content
                    'Dim val1 As String = Mid(strMyJson1, 3, 13)
                    'Return strMyJson1

                    If strMyJson1.Length < 100 Then Return "Error:" & strMyJson1
                    strMyJson1 = strMyJson1.Remove(2, 12)
                    Dim sqlquery As String = ""
                    Try
                        Dim dt As DataTable = ConvertJSONToDataTable(strMyJson1)
                        If dt.Columns.Count = 0 Then Return "Error: " & Response.Content
                        For i As Integer = 0 To dt.Rows.Count - 1
                            If dt.Rows(i).Item(0).ToString <> "" Then
                                Dim refid As String = Mid(dt.Rows(i).Item(0).ToString, 13)
                                sqlquery = "update [ARM FUND ACCOUNTS$Projected RSA Balance] set Remark = '" & dt.Rows(0).Item(1).ToString & "' where TransferRef = '" & refid & "'"
                                returndataset(sqlquery)
                            End If
                        Next

                    Catch ex As Exception
                        Return ex.Message & sqlquery
                    End Try
                    Return "SUCCESS" & sqlquery

                Else
                    Return "ERROR: " & Response.ErrorMessage & "-----" & json
                End If
                Return Response.Content.ToString
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function

        Public Function DownloadBalances(baseUrl As String, path As String, recordheader As ProvBalheader) As String
            'Dim record = New Record()
            Try


                Dim client As RestClient = New RestClient(baseUrl + path)
                Dim jsonheader = JsonConvert.SerializeObject(recordheader)
                '  Dim json = JsonConvert.SerializeObject(obj)
                Dim request = New RestRequest("", Method.GET)
                request.AddHeader("token", recordheader.token)
                request.AddHeader("userid", recordheader.userid)
                request.AddHeader("password", recordheader.password)
                request.AddHeader("quarter", recordheader.quarter)

                ' request.AddParameter("application/json; charset=utf-8", json, ParameterType.RequestBody)
                request.RequestFormat = DataFormat.Json


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
                Dim Response As IRestResponse = client.Execute(request)
                '   Return "AFTER EXECUTION STAGE" & Response.IsSuccessful & Response.ErrorMessage & "-----" & json
                If Response.IsSuccessful = True Then
                    Dim strMyJson1 As String = Response.Content
                    If strMyJson1.Length < 40 Then Return "Error:" & strMyJson1
                    Dim sqlquery As String = ""
                    Try
                        strMyJson1 = strMyJson1.Remove(2, 12)
                        Dim dt As DataTable = ConvertJSONToDataTable(strMyJson1)
                        If dt.Columns.Count = 0 Then Return "Error: " & Response.Content

                        strMyJson1 = Response.Content

                        strMyJson1 = strMyJson1.Remove(0, 36)
                        dt = ConvertJSONToDataTable(strMyJson1)
                        Dim ds As New DataSet
                        For i As Integer = 0 To dt.Rows.Count - 1

                            If dt.Rows(i).Item(0).ToString <> "" Then
                                Dim refid As String = dt.Rows(i).Item(0).ToString
                                ds = returndataset("select * from [ARM FUND ACCOUNTS$Projected RSA Balance] where TransferRef = '" & refid & "'")
                                If dt.Rows(i).Item(3).ToString = "" Then dt.Rows(i).Item(3) = "01-JAN-1753"
                                If ds.Tables(0).Rows.Count = 0 Then
                                    sqlquery = "INSERT INTO [ARM FUND ACCOUNTS$Projected RSA Balance]([TransferRef],[RSA_BALANCE],[AS_AT_DATE],[As_At_Date_String] ,[Remark] ,[QuaterId] ,[RSAPIN] ,[SURNAME] ,[MIDDLENAME] ,[FIRSTNAME] ,[TPFACODE])VALUES
 ('" & dt.Rows(i).Item(0).ToString & "','" & dt.Rows(i).Item(4).ToString & "' ,'" & CDate(dt.Rows(i).Item(3).ToString) & "','" & (dt.Rows(i).Item(3).ToString) & "','','" & (dt.Rows(i).Item(1).ToString) & "','" & (dt.Rows(i).Item(2).ToString) & "', '" & (dt.Rows(i).Item(5).ToString) & "','" & (dt.Rows(i).Item(8).ToString) & "', '" & (dt.Rows(i).Item(7).ToString) & "','" & (dt.Rows(i).Item(6).ToString) & "')"
                                Else

                                    sqlquery = "Update [ARM FUND ACCOUNTS$Projected RSA Balance] set [RSA_BALANCE] = '" & dt.Rows(i).Item(4).ToString & "',[AS_AT_DATE]= '" & CDate(dt.Rows(i).Item(3).ToString) & "',[As_At_Date_String]='" & dt.Rows(i).Item(3).ToString & "' ,[QuaterId] ='" & dt.Rows(i).Item(1).ToString & "'  where TransferRef = '" & refid & "'"

                                End If

                                ' sqlquery = "update [ARM FUND ACCOUNTS$Projected RSA Balances] set remark = '" & dt.Rows(0).Item(1).ToString & "' where TransferRef = '" & refid & "'"


                                returndataset(sqlquery)
                            End If
                        Next

                    Catch ex As Exception
                        Return ex.Message & sqlquery
                    End Try
                    Return "SUCCESS"

                Else
                    Return "ERROR: " & Response.ErrorMessage
                End If
                Return Response.Content.ToString
            Catch ex As Exception
                Return ex.Message
            End Try
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
        Public Function DownloadTH(baseUrl As String, path As String, recordheader As ProvBalheader, ByVal refids As String) As String
            'Dim record = New Record()
            Try


                Dim client As RestClient = New RestClient(baseUrl + path)
                Dim jsonheader = JsonConvert.SerializeObject(recordheader)
                '  Dim json = JsonConvert.SerializeObject(obj)
                Dim request = New RestRequest("", Method.POST)
                request.AddHeader("token", recordheader.token)
                request.AddHeader("userid", recordheader.userid)
                request.AddHeader("password", recordheader.password)
                request.AddHeader("quarter", recordheader.quarter)
                request.AddHeader("transferRefId", refids)
                ' request.AddParameter("application/json; charset=utf-8", json, ParameterType.RequestBody)
                request.RequestFormat = DataFormat.Json


                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
                Dim Response As IRestResponse = client.Execute(request)
                '   Return "AFTER EXECUTION STAGE" & Response.IsSuccessful & Response.ErrorMessage & "-----" & json
                If Response.IsSuccessful = True Then
                    Dim strMyJson1 As String = Response.Content

                    If strMyJson1.Length < 100 Then Return "Error: " & Response.Content
                    strMyJson1 = strMyJson1.Remove(2, 15)

                    'Dim val1 As String = Mid(strMyJson1, 3, 13)
                    Dim obj As Object = Split(strMyJson1, "detailRecords")
                    Dim detailsstring As String = Mid(obj(1).ToString, 3)
                    Dim detailhead As String = obj(0).ToString
                    'Dim myDataSet As DataTable = JsonConvert.DeserializeObject(Of DataTable)(detailsstring)
                    'Return strMyJson1
                    Dim sqlquery As String = ""
                    Try
                        Dim dt As DataTable = ConvertJSONToDataTable(detailsstring)
                        Dim dtheader As DataTable = ConvertJSONToDataTable(detailhead)
                        dt.AsEnumerable().Where(Function(row) row.ItemArray.All(Function(field) field Is Nothing Or field Is DBNull.Value Or field.Equals(""))).ToList().ForEach(Sub(row) row.Delete())
                        dt.AcceptChanges()
                        If dt.Columns.Count = 0 Then Return "Error: " & Response.Content
                        'Dim namestring(dt.Columns.Count - 1) As String
                        'For i As Integer = 0 To dt.Columns.Count - 1
                        '    namestring(i) = dt.Columns(i).ColumnName
                        'Next

                        dt = changedttype(dt)
                        'For i As Integer = 0 To dt.Columns.Count - 1
                        '    MsgBox(dt.Columns(i).DataType.Name)
                        'Next
                        '        Dim ndt As DataTable = ConvertToDataTable(Of DataTable)(namestring)
                        strMyJson1 = Response.Content

                        strMyJson1 = strMyJson1.Remove(0, 36)
                        'dt = ConvertJSONToDataTable(strMyJson1)
                        Dim ds As New DataSet
                        For i As Integer = 0 To dtheader.Rows.Count - 1

                            If dtheader.Rows(i).Item(0).ToString <> "" Then
                                Dim refid As String = refids
                                ds = returndataset("select * from [ARM FUND ACCOUNTS$Transaction History Header] where TH_Code = '" & refid & "'")
                                If ds.Tables(0).Rows.Count > 0 Then
                                    ds = returndataset("delete [ARM FUND ACCOUNTS$Transaction History Header]  where TH_Code = '" & refid & "'  delete [ARM FUND ACCOUNTS$Transaction History Lines]  where THH_Headder = '" & refid & "'")
                                End If
                                Dim empcode As String
                                empcode = Split(dtheader.Rows(0).Item(0).ToString, ":")(1).ToString
                                sqlquery = "INSERT INTO [ARM FUND ACCOUNTS$Transaction History Header] ([TH_Code]
,[TP_Code]  ,[Surname],[Firstname],[MiddlleName],[PIN],[Employer_Code],[Fund_Code],[Fund_Unit_Price]
,[Total_Fund_Unit],[RSA_Balance],[RSA_Gain_Loss],[Import_Date],[Imported_By],[Status])VALUES
 ('" & refid & "','','" & dtheader.Rows(0).Item("Surname").ToString & "'
,'" & dtheader.Rows(0).Item("Firstname").ToString & "','" & "" & "'
,'" & dtheader.Rows(0).Item("rsapin").ToString & "','" & empcode & "','" & dtheader.Rows(0).Item("FundCode").ToString & "'
,'" & dtheader.Rows(0).Item("UnitPrice").ToString & "','" & dtheader.Rows(0).Item("ttlNoofUnits").ToString & "'
,'" & dtheader.Rows(0).Item("RSABalance").ToString & "','" & dtheader.Rows(0).Item("ttlGainorLoss").ToString & "'
,'" & DateTime.Now & "','" & recordheader.userid & "','0')"

                                returndataset(sqlquery)
                            End If
                        Next


                        Using bulkCopy As SqlBulkCopy =
               New SqlBulkCopy(My.MySettings.Default.ARM_TESTINGConnectionString)

                            bulkCopy.DestinationTableName = "[ARM FUND ACCOUNTS$Transaction History Lines]"

                            Try
                                bulkCopy.ColumnMappings.Add("referenceId", "THH_Headder")
                                bulkCopy.ColumnMappings.Add("serialNo", "Line No")
                                bulkCopy.ColumnMappings.Add("emplContribution", "Employer_Contribution")
                                bulkCopy.ColumnMappings.Add("employeeContribution", "Employee_Contribution")
                                bulkCopy.ColumnMappings.Add("Fees", "Fees")
                                bulkCopy.ColumnMappings.Add("netContribution", "Net_Contributions")
                                bulkCopy.ColumnMappings.Add("numberOfUnits", "Number_Of_units")
                                bulkCopy.ColumnMappings.Add("others", "Other_Inflows")
                                bulkCopy.ColumnMappings.Add("paymentDate", "Pay_Receive_Date")
                                bulkCopy.ColumnMappings.Add("relatedMnthEnd", "RelatedMonthEnd")
                                bulkCopy.ColumnMappings.Add("relatedMnthStart", "RelatedMonthStart")
                                bulkCopy.ColumnMappings.Add("relatedPfaCode", "Related_PFA_Code")
                                bulkCopy.ColumnMappings.Add("totalContribution", "Total_Contribution")
                                bulkCopy.ColumnMappings.Add("transactionType", "Transaction_Type")
                                bulkCopy.ColumnMappings.Add("voluntaryContingent", "Voluntary_Contigent")
                                bulkCopy.ColumnMappings.Add("voluntaryRetirement", "Voluntary_Retirement")
                                bulkCopy.ColumnMappings.Add("withdrawal", "Other_Withdrawals")
                                bulkCopy.WriteToServer(dt)
                            Catch ex As Exception
                                Return ("THH LINES " & ex.Message)
                            End Try
                        End Using


                    Catch ex As Exception
                        Return ex.Message & sqlquery
                    End Try
                    Return "SUCCESS"

                Else
                    Return "ERROR: " & Response.ErrorMessage
                End If
                Return Response.Content.ToString
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function
        Private Function changedttype(ByVal dt As DataTable) As DataTable
            Dim dtCloned As DataTable = dt.Clone()
            dtCloned.Columns(0).DataType = GetType(Decimal)
            dtCloned.Columns(1).DataType = GetType(Decimal)
            dtCloned.Columns(2).DataType = GetType(Decimal)
            dtCloned.Columns(3).DataType = GetType(Decimal)
            dtCloned.Columns(4).DataType = GetType(Decimal)
            dtCloned.Columns(5).DataType = GetType(Decimal)
            dtCloned.Columns(6).DataType = GetType(Date)
            dtCloned.Columns(9).DataType = GetType(Date)
            dtCloned.Columns(10).DataType = GetType(Date)
            dtCloned.Columns(12).DataType = GetType(Integer)
            dtCloned.Columns(13).DataType = GetType(Decimal)
            dtCloned.Columns(15).DataType = GetType(Decimal)
            dtCloned.Columns(16).DataType = GetType(Decimal)
            dtCloned.Columns(17).DataType = GetType(Decimal)
            For Each row As DataRow In dt.Rows
                dtCloned.ImportRow(row)
            Next
            'For i As Integer = 0 To dt.Columns.Count - 1
            '    MsgBox(dtCloned.Columns(i).DataType.Name)
            'Next
            Return dtCloned
        End Function
        Public Function SubmitTH(baseUrl As String, path As String, refid As String, recordheader As ProvBalheader, obj As thObject) As String
            'Dim record = New Record()
            Try

                Dim client As RestClient = New RestClient(baseUrl + path)
                Dim jsonheader = JsonConvert.SerializeObject(recordheader)
                Dim json = JsonConvert.SerializeObject(obj)
                Dim request = New RestRequest("", Method.POST)
                request.AddHeader("token", recordheader.token)
                request.AddHeader("userid", recordheader.userid)
                request.AddHeader("password", recordheader.password)
                request.AddHeader("quarter", recordheader.quarter)
                request.AddHeader("transferRefId", refid)

                request.AddParameter("application/json; charset=utf-8", json, ParameterType.RequestBody)
                request.RequestFormat = DataFormat.Json


                Try
                    Dim serialWriter As StreamWriter
                    serialWriter = New StreamWriter("C:\Sendingfiles\" & refid & "TW.xml")
                    Dim xmlWriter As New XmlSerializer(request.Parameters.GetType())
                    xmlWriter.Serialize(serialWriter, request.Parameters)
                    serialWriter.Close()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 Or SecurityProtocolType.Tls11 Or SecurityProtocolType.Tls Or SecurityProtocolType.Ssl3
                Dim Response As IRestResponse = client.Execute(request)
                '   Return "AFTER EXECUTION STAGE" & Response.IsSuccessful & Response.ErrorMessage & "-----" & json
                If Response.IsSuccessful = True Then
                    Dim strMyJson1 As String = Response.Content
                    If strMyJson1.ToUpper.Contains("SYNTAX ERROR") Then Return "Error: " & Response.Content
                    'Dim val1 As String = Mid(strMyJson1, 3, 13)
                    'Return strMyJson1
                    strMyJson1 = strMyJson1.Remove(2, 12)
                    Dim sqlquery As String = ""
                    Try
                        Dim dt As DataTable = ConvertJSONToDataTable(strMyJson1)
                        If dt.Columns.Count = 0 Then Return "Error: " & Response.Content
                        If dt.Rows(0).Item(0).ToString.ToUpper.Contains("JAVA") Then Return "Error: " & Response.Content

                        For i As Integer = 0 To dt.Rows.Count - 1
                            If dt.Rows(i).Item(0).ToString <> "" Then
                                'Dim refid As String = Mid(dt.Rows(i).Item(0).ToString, 13)
                                'sqlquery = "update [ARM FUND ACCOUNTS$Projected RSA Balance] set Remark = '" & dt.Rows(0).Item(1).ToString & "' where TransferRef = '" & refid & "'"
                                'eftsharedservice.returndataset(sqlquery)
                            End If
                        Next

                    Catch ex As Exception
                        Return ex.Message & sqlquery
                    End Try
                    Return "SUCCESS" & sqlquery

                Else
                    Return "ERROR: " & Response.ErrorMessage & "-----" & json
                End If
                Return Response.Content.ToString
            Catch ex As Exception
                Return ex.Message
            End Try
        End Function

        Private Function ConvertJSONToDataTable(jsonString As String) As DataTable
            Dim dt As New DataTable
            'strip out bad characters
            Dim jsonParts As String() = jsonString.Replace("[", "").Replace("]", "").Split("},{")

            'hold column names
            Dim dtColumns As New List(Of String)

            'get columns
            For Each jp As String In jsonParts
                'only loop thru once to get column names
                Dim propData As String() = jp.Replace("{", "").Replace("}", "").Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                For Each rowData As String In propData
                    Try
                        Dim idx As Integer = rowData.IndexOf(":")
                        Dim n As String = rowData.Substring(0, idx - 1)
                        Dim v As String = rowData.Substring(idx + 1)
                        If Not dtColumns.Contains(n) Then
                            dtColumns.Add(n.Replace("""", ""))
                        End If
                    Catch ex As Exception
                        'Throw New Exception(String.Format("Error Parsing Column Name : {0}", rowData))
                    End Try

                Next
                Exit For
            Next

            'build dt
            For Each c As String In dtColumns
                dt.Columns.Add(c)
            Next
            'get table data
            For Each jp As String In jsonParts
                Dim propData As String() = jp.Replace("{", "").Replace("}", "").Split(New Char() {","}, StringSplitOptions.RemoveEmptyEntries)
                Dim nr As DataRow = dt.NewRow
                For Each rowData As String In propData
                    Try
                        Dim idx As Integer = rowData.IndexOf(":")
                        Dim n As String = rowData.Substring(0, idx - 1).Replace("""", "")
                        Dim v As String = rowData.Substring(idx + 1).Replace("""", "")
                        nr(n) = v
                    Catch ex As Exception
                        Continue For
                    End Try

                Next
                dt.Rows.Add(nr)
            Next
            Return dt
        End Function

    End Class

End Namespace