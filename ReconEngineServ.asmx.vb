Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Threading.Thread

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class ReconEngineServ
    Inherits System.Web.Services.WebService
    Public ReconNo, userid, attachuser, muploadBatchNo As String

    <WebMethod()>
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function
    Public fulldsctslines As New DataTable

    <WebMethod()>
    Public Function ReconEngines(ByVal headerno As String) As String
        Try

            Dim fqueries = " DELETE [RECONDB].[dbo].[Recon Engine Lines Temp] WHERE [No] ='" & headerno & "'"
            returndataset("RECONDB", fqueries)

            Dim recEngineLinestring As String = ""

            Dim db As New DataClasses1DataContext
            Dim db1 As New accountingdDataContext
            Dim dbrecon As New reconlinesclassDataContext

            Dim reconheader = (From i In db.AIICO_FUND_ACCOUNTS_Reconciliation_Engines Where i.No = headerno Select i.No, i.Fund, i.Fund_Type, i.Type_of_Remittance,
           i.Schedule_Monitor_No, i.Contribution_Period, i.Contribution_Type, i.CP_Header_No, i.Date, i.Status, i.Sub_Type, i.Reconciliation_DB, i.Schedule_Downloaded, i.Schedule_ID).FirstOrDefault

            ' Dim reconheader = (From i In db.AIICO_FUND_ACCOUNTS_Reconciliation_Engines Where i.No = headerno).FirstOrDefault

            If reconheader Is Nothing Then
                Return "Error Connecting to recon engine db"
            End If

            If reconheader.Status <> 0 Then
                Return "Recon Engine No: " + headerno + " not in the desired stage "
            End If

            Dim reconLines = (From i In dbrecon.Recon_Engine_Lines Where i.No = headerno)
            Dim reconlineTable As New DataTable
            reconlineTable = ToDataTables(reconLines)

            If reconLines Is Nothing Then
                Return "No line records"
            End If



            Dim linecount As Integer = reconLines.Count


            Dim pinstring As String = ""



            Dim dsvendor As New DataSet
            Dim query = "select No_ ,P_I_N ,Surname ,[First Name] ,[Middle Names], [Blocked] from [AIICO FUND ACCOUNTS$Vendor] where [P_I_N] in ("
            For i As Integer = 0 To linecount - 1
                pinstring = pinstring & "'" & reconlineTable.Rows(i).Item("PIN") & "',"
            Next

            pinstring = Mid(pinstring, 1, pinstring.Length - 1)

            pinstring = pinstring & ")"
            Dim vendords As DataSet
            query = query & pinstring
            vendords = returndataset("", query)

            If vendords Is Nothing Then
                Return "Error: Couldnt access the database at the moment..."
            End If
            Dim pery = reconheader.Contribution_Period
            '   Dim Q2 As String = " Select   distinct TOP " & vendords.Tables(0).Rows.Count & " * from [Recon Engine Lines] where PIN IN (" & pinstring & " and Period <>'" & pery & "' and [No] <> '" & headerno & "'  ORDER BY Period DESC"
            Dim Q2 As String = " Select   distinct TOP " & vendords.Tables(0).Rows.Count & " * from [Recon Engine Lines] where PIN IN (" & pinstring & "  and [No] <> '" & headerno & "'  ORDER BY Period DESC"
            Dim lineds As DataSet = returndataset("RECONDB", Q2)


            Dim vendor
            '   Dim vendorlist = (From k In db1.AIICO_FUND_ACCOUNTS_Vendors Where list.Contains(k.P_I_N) Select k.No_, k.First_Name, k.Surname, k.Middle Names, k.P_I_N)

            ' For Each item As Recon_Engine_Line In reconLines
            For i As Integer = 0 To reconlineTable.Rows.Count - 1
                Try
                    vendor = (From K In vendords.Tables(0).AsEnumerable
                              Where K.Field(Of String)("P_I_N").Contains(reconlineTable.Rows(i).Item("PIN"))).FirstOrDefault


                    '   If reconlineTable.Rows(i).Item("Line_No") = 144119 Then
                    '  MsgBox("here")
                    ' End If

                    If vendor Is Nothing Then

                        reconlineTable.Rows(i).Item("Not_on_NAV_DB") = True
                        reconlineTable.Rows(i).Item("Status") = 2 ' not in db
                        reconlineTable.Rows(i).Item("Accepted_Name") = ""
                        reconlineTable.Rows(i).Item("Suggested_employee_Name") = ""
                        'reconlineTable.Rows(i).Item("Suggested_First_Name") = ""
                        'reconlineTable.Rows(i).Item("Suggested_Middle_Name") = ""
                        reconlineTable.Rows(i).Item("Apply_Suggestions") = False
                        ' Return "" '"PIN DOESNT EXIST IN THE DB"
                    Else
                        reconlineTable.Rows(i).Item("Status") = 0
                        reconlineTable.Rows(i).Item("Accepted_Name") = ""
                        reconlineTable.Rows(i).Item("Suggested_employee_Name") = ""
                        'reconlineTable.Rows(i).Item("Suggested_First_Name") = ""
                        'reconlineTable.Rows(i).Item("Suggested_Middle_Name") = ""
                        reconlineTable.Rows(i).Item("Apply_Suggestions") = False



                        'Dim sur, sur1 As String
                        If UCase(reconlineTable.Rows(i).Item("Employee_Name").ToString.Trim).Contains(UCase(vendor.item("Surname").ToString.Trim)) = False Then
                            reconlineTable.Rows(i).Item("Suggested_Employee_Name") = vendor.item("Surname").ToString.Trim + " " + UCase(vendor.item("First Name").ToString.Trim) + " " + UCase(vendor.item("Middle Names").ToString.Trim)

                        End If
                        If UCase(reconlineTable.Rows(i).Item("Employee_Name").ToString.Trim).Contains(UCase(vendor.item("First Name").ToString.Trim)) = False Then
                            reconlineTable.Rows(i).Item("Suggested_Employee_Name") = vendor.item("Surname").ToString.Trim + " " + UCase(vendor.item("First Name").ToString.Trim) + " " + UCase(vendor.item("Middle Names").ToString.Trim)
                        End If


                        If UCase(reconlineTable.Rows(i).Item("Employee_Name").ToString.Trim).Contains(UCase(vendor.item("Middle Names").ToString.Trim)) = False Then
                            reconlineTable.Rows(i).Item("Suggested_Employee_Name") = vendor.item("Surname").ToString.Trim + " " + UCase(vendor.item("First Name").ToString.Trim) + " " + UCase(vendor.item("Middle Names").ToString.Trim)
                        End If


                        If reconlineTable.Rows(i).Item("Suggested_Employee_Name") = "" Then
                            reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                            'reconlineTable.Rows(i).Item("Accepted_Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First_Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim)
                            reconlineTable.Rows(i).Item("Accepted_Name") = reconlineTable.Rows(i).Item("Employee_Name")
                        End If

                        'If (UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim) = UCase(vendor.item("Surname").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("First_Name").ToString.Trim) = UCase(vendor.item("First Name").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim) = UCase(vendor.item("Middle Names").ToString.Trim)) Then

                        '    reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                        '    reconlineTable.Rows(i).Item("Accepted_Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First_Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim)
                        'End If



                        'WRONG PERIOD
                        Dim pin, per As String
                        pin = reconlineTable.Rows(i).Item("PIN")
                        If IsDBNull(reconlineTable.Rows(i).Item("period")) Then
                            reconlineTable.Rows(i).Item("period") = reconheader.Contribution_Period
                        End If

                        If reconlineTable.Rows(i).Item("period").ToString = "" Then
                            reconlineTable.Rows(i).Item("period") = reconheader.Contribution_Period
                        End If
                        per = reconlineTable.Rows(i).Item("period")

                        If per.ToString.Trim = "" Then per = reconheader.Contribution_Period





                        Dim liness
                        liness = (From K In lineds.Tables(0).AsEnumerable
                                  Where K.Field(Of String)("PIN").Contains(pin)).FirstOrDefault


                        If liness Is Nothing = False Then
                            If liness.Item("Period").ToString = pery Then
                                reconlineTable.Rows(i).Item("Wrong_Period") = True
                                reconlineTable.Rows(i).Item("Status") = 4  'wrong period
                            End If
                            'PREVIOUS AMOUNT
                            'type of remittance = monthly
                            If reconheader.Type_of_Remittance = 0 Then
                                reconlineTable.Rows(i).Item("PREVIOUS_AMOUNT") = liness.Item("Total Contribution").ToString
                            End If
                        Else
                            reconlineTable.Rows(i).Item("PREVIOUS_AMOUNT") = 0
                        End If
                    End If

                    'FOR BLOCKED PIN
                    Try
                        If vendor.item("Blocked") = 2 Then
                            reconlineTable.Rows(i).Item("Blocked_PIN") = 1
                            reconlineTable.Rows(i).Item("Status") = 5
                        End If
                    Catch ex As Exception

                    End Try


                    Dim duplic = (From K In reconlineTable.AsEnumerable
                                  Where K.Field(Of String)("PIN").Contains(reconlineTable.Rows(i).Item("PIN"))).Count
                    If duplic > 1 Then
                        reconlineTable.Rows(i).Item("Status") = 3
                        reconlineTable.Rows(i).Item("Wrong_Period") = False
                        reconlineTable.Rows(i).Item("Duplicate_Record") = True
                        'duplicate
                    End If


                Catch ex As Exception

                End Try

            Next


            reconlineTable.AcceptChanges()
            '  dbrecon.SubmitChanges()
            Dim queries As String = ""
            queries = " DELETE [dbo].[Recon Engine Lines Temp] WHERE [No] ='" & headerno & "'"
            returndataset("RECONDB", queries)

            ' db.SubmitChanges()

            Dim con1 As New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            con1.Open()

            Using bulkCopy As SqlBulkCopy =
              New SqlBulkCopy(con1)
                bulkCopy.DestinationTableName = "RECONDB.dbo.[Recon Engine Lines Temp]"

                Try
                    ' Write from the source to the destination.
                    bulkCopy.WriteToServer(reconlineTable)

                Catch ex As Exception
                    con1.Close()
                    Return (ex.Message)

                End Try
                ' con1.Close()
            End Using
            'Return "error testing here"

            queries = ""
            queries += "    merge into [RECONDB].[dbo].[Recon Engine Lines] as Target"
            queries += " Using [RECONDB].[dbo].[Recon Engine Lines Temp] As Source"
            queries += " On Target.[No]=Source.[No]  And Target.[Line No]=Source.[Line No]"
            queries += " When matched Then "
            queries += " update set Target.[Suggested Employee Name]=Source.[Suggested Employee Name],"
            queries += " Target.[Status] = Source.[Status], Target.[PREVIOUS AMOUNT] = Source.[PREVIOUS AMOUNT], Target.[Wrong Period] = Source.[Wrong Period], Target.[Duplicate Record] = Source.[Duplicate Record],  Target.[Accepted Name] = Source.[Accepted Name],Target.[Blocked PIN] = Source.[Blocked PIN],"
            queries += " Target.[Not on NAV DB] = Source.[Not on NAV DB];"

            Dim com As New SqlCommand
            com = New SqlCommand(queries, con1)
            If com.ExecuteNonQuery <> 0 Then
                con1.Close()
                queries = " DELETE [RECONDB].[dbo].[Recon Engine Lines Temp] WHERE [No] ='" & headerno & "'"
                queries += " update AIICO.dbo.[AIICO FUND ACCOUNTS$Reconciliation Engine] set [Status] = 1, [Reconciliation DB] =1 where No='" & headerno & "'"

                returndataset("RECONDB", queries)

                con1.Dispose()
                'reconheader.Status = 1
                'reconheader.Reconciliation_DB = 1

                db.SubmitChanges()
            Else

            End If


            Return ""
        Catch ex As Exception
            Return "Error " & ex.Message
        End Try
    End Function
    Public Function ToDataTables(Of T)(ByVal varlist As IEnumerable(Of T), Optional ByVal B As String = "") As DataTable
        Dim dtReturn As New DataTable()

        ' column names  
        Dim oProps As PropertyInfo() = Nothing

        If varlist Is Nothing Then
            Return dtReturn
        End If

        For Each rec As T In varlist
            ' Use reflection to get property names, to create table, Only first time, others will follow  
            If oProps Is Nothing Then
                oProps = DirectCast(rec.[GetType](), Type).GetProperties()
                For Each pi As PropertyInfo In oProps
                    Dim colType As Type = pi.PropertyType

                    If (colType.IsGenericType) AndAlso (colType.GetGenericTypeDefinition() Is GetType(Nullable(Of ))) Then
                        colType = colType.GetGenericArguments()(0)
                    End If

                    dtReturn.Columns.Add(New DataColumn(pi.Name, colType))
                Next
            End If

            Dim dr As DataRow = dtReturn.NewRow()

            For Each pi As PropertyInfo In oProps
                dr(pi.Name) = If(pi.GetValue(rec, Nothing) Is Nothing, DBNull.Value, pi.GetValue(rec, Nothing))
            Next
            dtReturn.Rows.Add(dr)
        Next

        oProps = Nothing
        Return dtReturn

    End Function
    Public exerror As String
    Private Function returndataset(ByVal databasenames As String, ByVal query As String) As DataSet
        returndataset = Nothing
        exerror = ""
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            ' con1.ConnectionString = "Data Source=192.168.0.30;Initial Catalog=ARM;User ID=sa;Password=t@sting1"
            con1.Open()
            If databasenames <> "" Then con1.ChangeDatabase(databasenames)
            com1 = New SqlCommand(query, con1)
            com1.CommandTimeout = 0
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

            exerror = ex.Message
            Return Nothing
        End Try
    End Function
    Private Function ctsAlreadyAssigned(ByVal uploadbatchno As String, ByVal dsct As DataSet)
        Dim fullnct As New DataTable
        fullnct = dsct.Tables(2).Clone
        Dim Query = "select * from AIICO.dbo.[AIICO FUND ACCOUNTS$UPCL CTS Manager] WHERE [Schedule Monitor ID]='" & uploadbatchno & "' and [Status] = 1 "
        Dim dsexist As DataSet = returndataset("ARM", Query)
        If dsexist.Tables(0).Rows.Count = 0 Then
            Return "Error: No Schedule imported with this batch no."
        End If
        fullnct.Rows.Clear()
        For i As Integer = 0 To dsexist.Tables(0).Rows.Count - 1
            Dim fullctslinesexist As DataTable = dsct.Tables(2)
            Dim reconquery As String = "select * from [Recon Engine Lines] where No='" & dsexist.Tables(0).Rows(i).Item("Recon Batch No").ToString & "'"
            Dim linesrespds = returndataset("Recondb", reconquery)
            If linesrespds.Tables(0).Rows.Count > 0 Then
                Dim scheduleidnew As String = dsexist.Tables(0).Rows(i).Item("Schedule ID").ToString
                Dim transidnew As String = dsexist.Tables(0).Rows(i).Item("trans ID").ToString
                Try
                    Dim dsctslinesnew = (From l In fullctslinesexist.AsEnumerable
                                         Where l.Field(Of String)("SCHEDULE ID").Contains(scheduleidnew) And l.Field(Of String)("trans ID").Contains(transidnew)).CopyToDataTable()

                    For s = 0 To dsctslinesnew.Rows.Count - 1
                        dsctslinesnew.Rows(s).Item("Assigned By") = dsexist.Tables(0).Rows(i).Item("Uploaded By").ToString
                        dsctslinesnew.Rows(s).Item("Assigned") = 1
                        dsctslinesnew.Rows(s).Item("Assigned Date") = DateTime.Now
                        dsctslinesnew.Rows(s).Item("Uploader Assigned") = dsexist.Tables(0).Rows(i).Item("Assigned To").ToString
                        dsctslinesnew.Rows(s).Item("Recon Batch No") = dsexist.Tables(0).Rows(i).Item("Recon Batch No").ToString
                        dsctslinesnew.AcceptChanges()
                        fullnct.ImportRow(dsctslinesnew.Rows(s))
                        fullnct.AcceptChanges()
                    Next
                    '     Return "Success"
                Catch ex As Exception
                    Return ex.Message
                End Try
            End If
        Next
        Dim ndt As New DataTable
        ndt = Nothing
        'Return fullnct.Rows.Count
        Return finalvalid(uploadbatchno, ndt, fullnct)

    End Function

    <WebMethod()>
    Public Function AssignCTSProcess(ByVal workdate As Date, ByVal uploadBatchNo As String, ByVal userid As String, ByVal withcorrection As Boolean) As String
        Dim availuserdt As New DataTable
        muploadBatchNo = uploadBatchNo
        Dim dsctsHeader As DataSet
        'Dim dsctsLines As DataSet
        Dim query As String
        'Dim reconNo As String
        ' Dim attachuser As String = ""
        Dim FNAME As String = ""
        Dim MNAME As String = ""
        Dim SNAME As String = ""
        Dim k As Integer = 0
        query = "select * from AIICO.dbo.[AIICO FUND ACCOUNTS$UPCL CTS Manager] WHERE [Schedule Monitor ID]='" & uploadBatchNo & "' and [Status] = 0  order by [TRANS ID]   select * from AIICO.dbo.[User] where [Recon Processor] = 1   select * from [UPCL CTS] WHERE [Schedule Monitor No]='" & uploadBatchNo & "' and Assigned = 0 select * from [Recon Engine Lines] where No='' "

        dsctsHeader = returndataset("recondb", query)
        If dsctsHeader.Tables.Count = 0 Then Return "Error: Connection Issue"
        If dsctsHeader.Tables(0).Rows.Count = 0 Then
            Return ctsAlreadyAssigned(uploadBatchNo, dsctsHeader)
        End If




        availuserdt = dsctsHeader.Tables(1)
        If availuserdt.Rows.Count = 0 Then Return "Error: No user available to assign CTS to"
        Dim fullctslines As DataTable = dsctsHeader.Tables(2)
        fulldsctslines = fullctslines.Clone
        Dim reconlineTable As DataTable = dsctsHeader.Tables(3)
        ' reconlineTable = dsctsHeader.Tables(3)
        Dim scheduleid As String
        Dim transid As String = ""
        attachuser = availuserdt.Rows(0).Item("User Name").ToString
        For i As Integer = 0 To dsctsHeader.Tables(0).Rows.Count - 1
            If k > availuserdt.Rows.Count - 1 Then k = 0
            If transid <> dsctsHeader.Tables(0).Rows(i).Item("TRANS ID").ToString Then
                transid = dsctsHeader.Tables(0).Rows(i).Item("TRANS ID").ToString
                attachuser = availuserdt.Rows(k).Item("User Name").ToString
                k += 1
            End If

            scheduleid = dsctsHeader.Tables(0).Rows(i).Item("Schedule ID").ToString



            transid = dsctsHeader.Tables(0).Rows(i).Item("trans ID").ToString
            'create reconengine header
            Dim db As New DataClasses1DataContext
            Dim dbrecon As New reconlinesclassDataContext
            Dim dsctslines As New DataTable
            Dim ReconEngineHeader As New AIICO_FUND_ACCOUNTS_Reconciliation_Engine
            ReconNo = getReconNo(ReconNo)


            ReconEngineHeader.No = ReconNo
            ReconEngineHeader.User_ID = attachuser
            ReconEngineHeader.Date = workdate
            ReconEngineHeader.Fund = "RSA"
            ReconEngineHeader.Value_Date = workdate
            ReconEngineHeader.Schedule_ID = scheduleid
            ReconEngineHeader.Contribution_Type = dsctsHeader.Tables(0).Rows(i).Item("Contribution Type")
            'ReconEngineHeader.Contribution_Type = 1
            ReconEngineHeader.No_Series = ""
            ReconEngineHeader.Reconciliation_DB = 0
            ReconEngineHeader.Schedule_Downloaded = 1
            ReconEngineHeader.Source = ""
            ReconEngineHeader.Status = 0
            ReconEngineHeader.Sub_Type = ""
            ReconEngineHeader.Type_of_Remittance = 0
            Dim contperiod As Date
            Dim contperiodstr As String = ""
            contperiod = dsctsHeader.Tables(0).Rows(i).Item("Schedule Date")
            Dim schedyear As Integer = contperiod.Year
            Dim schedmonth As Integer = contperiod.Month
            contperiodstr = (schedyear.ToString & "-" & schedmonth.ToString.PadLeft(2, "0") & "-" & "01")
            Date.TryParse(contperiodstr, contperiod)
            'contperiod = CDate(contperiodstr)

            ReconEngineHeader.Contribution_Period = contperiod
            If dsctsHeader.Tables(0).Rows(i).Item("Employer code").ToString = "0" Then dsctsHeader.Tables(0).Rows(i).Item("Employer code") = "000000000000"
            ReconEngineHeader.Employer_Code = dsctsHeader.Tables(0).Rows(i).Item("Employer code").ToString
            ReconEngineHeader.Employer_Name = dsctsHeader.Tables(0).Rows(i).Item("Employer Name").ToString
            ReconEngineHeader.Period_Narration = MonthName(contperiod.Month()) + " " + contperiod.Year.ToString
            ReconEngineHeader.Description = Mid(ReconEngineHeader.Employer_Name, 1, 40) + " " + ReconEngineHeader.Period_Narration
            ReconEngineHeader.CP_Header_No = ""
            ReconEngineHeader.Description = Mid(ReconEngineHeader.Description, 1, 50)
            ReconEngineHeader.Date = workdate

            ReconEngineHeader.Downloaded_By = userid
            ReconEngineHeader.Fund_Type = 1
            ReconEngineHeader.Select = 0
            ReconEngineHeader.Selected_by = ""
            ReconEngineHeader.Schedule_ID = scheduleid
            ReconEngineHeader.Datetime_Archive = "1753-01-01"
            ReconEngineHeader.Schedule_Monitor_No = muploadBatchNo
            ReconEngineHeader.Archived_by = ""

            dsctsHeader.Tables(0).Rows(i).Item("Recon Batch No") = ReconNo
            dsctsHeader.Tables(0).Rows(i).Item("Assigned To") = attachuser


            'create reconengine lines
            'dsctslines = New DataTable
            reconlineTable = New DataTable
            Try
                dsctslines = (From l In fullctslines.AsEnumerable
                              Where l.Field(Of String)("SCHEDULE ID").Contains(scheduleid) And l.Field(Of String)("TRANS ID").Contains(transid)).CopyToDataTable()
                reconlineTable = dsctsHeader.Tables(3)


            Catch ex As Exception

            End Try
            If dsctslines Is Nothing = False Then
                If dsctslines.Rows.Count > 0 Then
                    ManageeachScheduleid(dsctslines, reconlineTable)
                    db.AIICO_FUND_ACCOUNTS_Reconciliation_Engines.InsertOnSubmit(ReconEngineHeader)
                    db.SubmitChanges()
                End If
            End If
        Next

        Dim resval = finalvalid(uploadBatchNo, reconlineTable, fulldsctslines)
        If resval <> "" Then Return resval
        Dim queries As String
        queries = ""

        Dim querh As String = ""
        For i As Integer = 0 To dsctsHeader.Tables(0).Rows.Count - 1
            querh += "update AIICO.dbo.[AIICO FUND ACCOUNTS$UPCL CTS Manager] set [Recon Batch No] = '" & dsctsHeader.Tables(0).Rows(i).Item("Recon Batch No") & "', [Status] = 1, [Assigned To]='" & dsctsHeader.Tables(0).Rows(i).Item("Assigned To") & "' WHERE [Schedule Monitor ID]='" & uploadBatchNo & "' and [Schedule ID]= '" & dsctsHeader.Tables(0).Rows(i).Item("Schedule ID").ToString & "' and [TRANS ID]= '" & dsctsHeader.Tables(0).Rows(i).Item("trans ID").ToString & "'"
            If (i Mod 5 = 10) Or (i = dsctsHeader.Tables(0).Rows.Count - 1) Then
                returndataset("RECONDB", querh)
                querh = ""
            End If
        Next
        'do recon webservice 
        If withcorrection Then
            For i As Integer = 0 To dsctsHeader.Tables(0).Rows.Count - 1
                Dim nreconlinetb = New DataTable
                Try
                    Dim recHeader As String = dsctsHeader.Tables(0).Rows(i).Item("Recon Batch No")
                    nreconlinetb = (From l In reconlineTable.AsEnumerable
                                    Where l.Field(Of String)("No").Contains(recHeader)).CopyToDataTable()
                    ReconEnginesInternal(recHeader, nreconlinetb)
                Catch ex As Exception

                End Try
            Next
        End If
    End Function
    'Private Function ManageeachScheduleid(ByVal) As String
    Private Function ManageeachScheduleid(ByVal dsctslinesM As DataTable, ByVal reconlineTableM As DataTable) As String

        Dim reconID = ReconNo
        Dim attachedUser = attachuser
        Dim userAssigner = userid
        Dim UPLOADBATCHNOM = muploadBatchNo
        Dim dr As DataRow
        Dim fullname As String
        Dim namecount As Integer
        Dim schedDate As Date
        Dim schedyear, schedmonth As Integer
        Dim sname As String = ""
        Dim fname As String = ""
        Dim mName As String = ""
        If dsctslinesM Is Nothing Then Return ""
        Dim scheduleid As String = dsctslinesM.Rows(0).Item("SCHEDULE ID").ToString
        Dim transid As String = dsctslinesM.Rows(0).Item("trans ID").ToString
        For j As Integer = 0 To dsctslinesM.Rows.Count - 1
            fullname = dsctslinesM.Rows(j).Item("EMPLOYEE NAME").ToString.ToUpper.Trim
            fullname = fullname.Replace("MR.", "")
            fullname = fullname.Replace("MRS.", "")
            fullname = fullname.Replace("MS.", "")
            fullname = fullname.Replace("MR ", "")
            fullname = fullname.Replace("MRS ", "")
            fullname = fullname.Replace("MS ", "")
            fullname = fullname.Trim

            schedDate = dsctslinesM.Rows(j).Item("SCHEDULE DATE")
            schedyear = schedDate.Year
            schedmonth = schedDate.Month

            Date.TryParse(schedyear.ToString & "-" & schedmonth.ToString & "-" & "01", schedDate)
            fullname = RemoveCharacter(fullname)
            Dim obj As Object = Split(fullname.ToString, " ")
            namecount = UBound(obj)
            sname = obj(0).ToString.Trim
            mName = ""
            fname = ""
            If namecount > 0 Then fname = obj(1).ToString.Trim
            If namecount > 1 Then mName = obj(2).ToString.Trim

            dr = reconlineTableM.NewRow
            dr.Item("No") = ReconNo
            dr.Item("Line No") = j * 1000
            dr.Item("PIN") = Mid(dsctslinesM.Rows(j).Item("PIN").ToString, 1, 20)
            dr.Item(12) = 0
            dr.Item(13) = 0
            dr.Item(14) = 0
            dr.Item(15) = 0
            'dr.Item("Surname") = Mid(sname.ToString, 1, 50)
            'dr.Item("First Name") = Mid(fname.ToString, 1, 50)
            'dr.Item("Middle Name") = Mid(mName.ToString, 1, 50)
            dr.Item("Employee Name") = fullname
            dr.Item("Employee Contribution") = Math.Round(CDec(dsctslinesM.Rows(j).Item(13).ToString), 2)
            dr.Item("Employer Contribution") = Math.Round(CDec(dsctslinesM.Rows(j).Item(14).ToString), 2)
            dr.Item("Employee AVC") = Math.Round(CDec(dsctslinesM.Rows(j).Item(15).ToString), 2)
            dr.Item("Employer AVC") = Math.Round(CDec(dsctslinesM.Rows(j).Item(16).ToString), 2)
            dr.Item("Interest Penalty") = 0
            dr.Item("Total Contribution") = Math.Round(dr.Item(5) + dr.Item(6) + dr.Item(7) + dr.Item(8) + dr.Item(9), 2)
            dr.Item("Suggested employee Name") = ""
            'dr.Item("Suggested First Name") = ""
            'dr.Item("Suggested Middle Name") = ""
            dr.Item("Status") = 0
            dr.Item("Apply Suggestions") = 0
            dr.Item("Accepted Name") = ""
            dr.Item("Date of Update") = "1753-01-01"
            dr.Item("Time of Update") = "1753-01-01"
            dr.Item("Accepted By") = ""
            dr.Item("Not on NAV DB") = 0
            dr.Item("Duplicate Record") = 0
            dr.Item(23) = 0
            dr.Item(25) = "1753-01-01"
            dr.Item(31) = 0
            dr.Item(26) = 0
            dr.Item(27) = Mid(dsctslinesM.Rows(j).Item("CSV_Name").ToString, 1, 30)
            dr.Item("Period") = "1753-01-01"
            dr.Item(29) = Mid(scheduleid.ToString, 1, 30)
            dr.Item(30) = Mid(transid.ToString, 1, 30)
            dr.Item(31) = ""

            dr.Item(33) = muploadBatchNo  ' Mid(dsctslinesM.Rows(j).Item("CTS ID").ToString, 1, 40)
            reconlineTableM.Rows.Add(dr)

            dsctslinesM.Rows(j).Item("Assigned By") = userAssigner
            dsctslinesM.Rows(j).Item("Assigned") = 1
            dsctslinesM.Rows(j).Item("Assigned Date") = DateTime.Now
            dsctslinesM.Rows(j).Item("Uploader Assigned") = attachedUser
            dsctslinesM.Rows(j).Item("Recon Batch No") = reconID
            dsctslinesM.AcceptChanges()

            fulldsctslines.ImportRow(dsctslinesM.Rows(j))
            fulldsctslines.AcceptChanges()
        Next
    End Function
    Function RemoveCharacter(ByVal stringToCleanUp) As String
        Dim characterToRemove As String = ""
        characterToRemove = Chr(34) + "#$%()*+/\~"
        Dim firstThree As Char() = characterToRemove.Take(16).ToArray()
        For index = 1 To firstThree.Length - 1
            stringToCleanUp = stringToCleanUp.ToString.Replace(firstThree(index), "")
        Next
        Return stringToCleanUp
    End Function

    Private Function finalvalid(ByVal uploadbatch As String, ByVal reconlinetb As DataTable, ByVal ctslinetempdb As DataTable) As String
        If reconlinetb Is Nothing = False Then

            Using bulkCopy As SqlBulkCopy =
                 New SqlBulkCopy(My.MySettings.Default.ARM_TESTINGConnectionString)

                bulkCopy.DestinationTableName = "RECONDB.dbo.[Recon Engine Lines]"
                Try
                    bulkCopy.WriteToServer(reconlinetb)
                Catch ex As Exception
                    Return ("rec eng " & ex.Message)
                End Try
            End Using
        End If

        'update ctsmanager (status, assignedto,reconbatchno
        'dump update into ctslines 
        If ctslinetempdb Is Nothing = False Then
            Using bulkCopy As SqlBulkCopy =
             New SqlBulkCopy(My.MySettings.Default.ARM_TESTINGConnectionString)
                bulkCopy.DestinationTableName = "RECONDB.dbo.[UPCL CTS TEMP]"
                Try
                    bulkCopy.WriteToServer(ctslinetempdb)
                Catch ex As Exception
                    Return (ex.Message)
                End Try
            End Using
        End If

        Dim CQUERY As String = "WITH cte AS ( SELECT  [PIN],  [SCHEDULE ID], [Schedule Monitor No],  [trans ID], ROW_NUMBER() OVER (
            PARTITION BY    [SCHEDULE ID],   [Schedule Monitor No] ORDER BY   [PIN], [SCHEDULE ID],  [Schedule Monitor No] ) row_num FROM   [UPCL CTS TEMP])DELETE FROM cte WHERE row_num > 1;"
        returndataset("RECONDB", CQUERY)

        'update ctslines (assigned,uploaderassigned, assigned by,reconbatchno,assigned date,
        Dim queries = " merge into [RECONDB].[dbo].[UPCL CTS] As Target Using [RECONDB].[dbo].[UPCL CTS TEMP] As Source On
	 Target.[SCHEDULE ID]= Source.[SCHEDULE ID] AND Target.[Schedule Monitor No]='" & uploadbatch & "'   And Target.[trans ID]=source.[trans ID] And Target.[PIN]=Source.[PIN]
  When matched Then  
	 	 update set 
TARGET.[SCHEDULE DATE]= SOURCE.[SCHEDULE DATE],
TARGET.[VALUE DATE]= SOURCE.[VALUE DATE],
TARGET.[PAYMENT  DATE]= SOURCE.[PAYMENT  DATE],
TARGET.MONTH = SOURCE.MONTH,
TARGET.YEAR = SOURCE.YEAR,
TARGET.[EMPLOYER CODE]= SOURCE.[EMPLOYER CODE],
TARGET.[EMPLOYER NAME]= SOURCE.[EMPLOYER NAME],
TARGET.[FUND CODE RSA]= SOURCE.[FUND CODE RSA],
TARGET.[EMPLOYEE NAME]= SOURCE.[EMPLOYEE NAME],
TARGET.PIN = SOURCE.PIN,
TARGET.[EMPLOYEE CONTRIBUTION]= SOURCE.[EMPLOYEE CONTRIBUTION],
TARGET.[EMPLOYER CONTRIBUTION]= SOURCE.[EMPLOYER CONTRIBUTION],
TARGET.[EMPLOYEE VOL]= SOURCE.[EMPLOYEE VOL],
TARGET.[EMPLOYER VOL]= SOURCE.[EMPLOYER VOL],
TARGET.TOTAL = SOURCE.TOTAL,
TARGET.[SORT CODE]= SOURCE.[SORT CODE],TARGET.[CURRENCY CODE]= SOURCE.[CURRENCY CODE],
TARGET.[trans ID]= SOURCE.[trans ID],
TARGET.[CSV_NAME]= SOURCE.[CSV_NAME],
TARGET.Assigned = SOURCE.Assigned,
TARGET.[Uploader Assigned]= SOURCE.[Uploader Assigned],
TARGET.[Assigned By]= SOURCE.[Assigned By],
TARGET.[Recon Batch No]= SOURCE.[Recon Batch No],
TARGET.[Assigned Date]= SOURCE.[Assigned Date],
TARGET.[Schedule Monitor No]= SOURCE.[Schedule Monitor No];"

        returndataset("MASTER", queries)
        If exerror = "" Then

            queries = " DELETE [RECONDB].[dbo].[UPCL CTS TEMP] WHERE [Schedule Monitor No]='" & uploadbatch & "'"
            returndataset("RECONDB", queries)
        End If
        Return "" '"Previously Assigned"
    End Function

    Private Function getReconNo(ByVal oldrecno As String) As String
        If oldrecno = "" Then
            Dim query As String = "select TOP 1 [No] from AIICO.dbo.[AIICO FUND ACCOUNTS$Reconciliation Engine] order by No desc "
            Dim DS As New DataSet
            DS = returndataset("RECONDB", query)
            oldrecno = DS.Tables(0).Rows(0).Item(0).ToString
        End If
        oldrecno = oldrecno.Replace("R", "0")

        Dim longint As Integer
        longint = Convert.ToInt64(oldrecno)
        longint = longint + 1
        getReconNo = "R" + longint.ToString.PadLeft(19, "0")
        Return getReconNo
    End Function
    Public Function ReconEnginesInternal(ByVal headerno As String, ByVal reconlineTable As DataTable) As String
        Try
            Dim recEngineLinestring As String = ""

            Dim db As New DataClasses1DataContext
            Dim db1 As New accountingdDataContext
            Dim dbrecon As New reconlinesclassDataContext

            Dim reconheader = (From i In db.AIICO_FUND_ACCOUNTS_Reconciliation_Engines Where i.No = headerno).FirstOrDefault
            If reconheader Is Nothing Then
                Return "Error Connecting to recon engine db"
            End If

            If reconheader.Status <> 0 Then
                Return "Recon Engine No: " + headerno + " not in the desired stage "
            End If

            'Dim reconLines = (From i In dbrecon.Recon_Engine_Lines Where i.No = headerno)
            'Dim reconlineTable As New DataTable
            'reconlineTable = ToDataTables(reconLines)

            'If reconLines Is Nothing Then
            '    Return "No line records"
            'End If



            Dim linecount As Integer = reconlineTable.Rows.Count


            Dim pinstring As String = ""



            Dim dsvendor As New DataSet
            Dim query = "select No_ ,P_I_N ,Surname ,[First Name] ,[Middle Names], [Blocked] from [AIICO FUND ACCOUNTS$Vendor] where [P_I_N] in ("
            For i As Integer = 0 To linecount - 1
                pinstring = pinstring & "'" & reconlineTable.Rows(i).Item("PIN") & "',"
            Next

            pinstring = Mid(pinstring, 1, pinstring.Length - 1)

            pinstring = pinstring & ")"
            Dim vendords As DataSet
            query = query & pinstring
            vendords = returndataset("", query)

            If vendords Is Nothing Then
                Return "Error: Couldnt access the database at the moment..."
            End If
            Dim pery = reconheader.Contribution_Period
            '   Dim Q2 As String = " Select   distinct TOP " & vendords.Tables(0).Rows.Count & " * from [Recon Engine Lines] where PIN IN (" & pinstring & " and Period <>'" & pery & "' and [No] <> '" & headerno & "'  ORDER BY Period DESC"
            Dim Q2 As String = " Select   distinct TOP " & vendords.Tables(0).Rows.Count & " * from [Recon Engine Lines] where PIN IN (" & pinstring & "  and [No] <> '" & headerno & "'  ORDER BY Period DESC"
            Dim lineds As DataSet = returndataset("RECONDB", Q2)


            Dim vendor
            '   Dim vendorlist = (From k In db1.AIICO_FUND_ACCOUNTS_Vendors Where list.Contains(k.P_I_N) Select k.No_, k.First Name, k.Surname, k.Middle Names, k.P_I_N)

            ' For Each item As Recon_Engine_Line In reconLines
            For i As Integer = 0 To reconlineTable.Rows.Count - 1
                vendor = (From K In vendords.Tables(0).AsEnumerable
                          Where K.Field(Of String)("P_I_N").Contains(reconlineTable.Rows(i).Item("PIN"))).FirstOrDefault


                '   If reconlineTable.Rows(i).Item("Line_No") = 144119 Then
                '  MsgBox("here")
                ' End If

                If vendor Is Nothing Then

                    reconlineTable.Rows(i).Item("Not on NAV DB") = True
                    reconlineTable.Rows(i).Item("Status") = 2 ' not in db
                    reconlineTable.Rows(i).Item("Accepted Name") = ""
                    reconlineTable.Rows(i).Item("Suggested Surname") = ""
                    reconlineTable.Rows(i).Item("Suggested First Name") = ""
                    reconlineTable.Rows(i).Item("Suggested Middle Name") = ""
                    reconlineTable.Rows(i).Item("Apply Suggestions") = False
                    ' Return "" '"PIN DOESNT EXIST IN THE DB"
                Else
                    reconlineTable.Rows(i).Item("Status") = 0
                    reconlineTable.Rows(i).Item("Accepted Name") = ""
                    reconlineTable.Rows(i).Item("Suggested Surname") = ""
                    reconlineTable.Rows(i).Item("Suggested First Name") = ""
                    reconlineTable.Rows(i).Item("Suggested Middle Name") = ""
                    reconlineTable.Rows(i).Item("Apply Suggestions") = False



                    'Dim sur, sur1 As String
                    If UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim) <> UCase(vendor.item("Surname").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested Surname") = vendor.item("Surname").ToString.Trim

                    End If
                    If UCase(reconlineTable.Rows(i).Item("First Name").ToString.Trim) <> UCase(vendor.item("First Name").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested First Name") = vendor.item("First Name").ToString.Trim
                    End If
                    If UCase(reconlineTable.Rows(i).Item("Middle Name").ToString.Trim) <> UCase(vendor.item("Middle Names").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested Middle Name") = vendor.item("Middle Names").ToString()
                    End If




                    'If (UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim) = UCase(vendor.item("Surname").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("First Name").ToString.Trim) = UCase(vendor.item("First Name").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("Middle Name").ToString.Trim) = UCase(vendor.item("Middle Names").ToString.Trim)) Then

                    '    reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                    '    reconlineTable.Rows(i).Item("Accepted Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle Name").ToString.Trim)
                    'End If
                    Dim surname As String = ""
                    surname = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim)

                    If UCase(vendor.item("Surname").ToString.Trim.Contains(surname)) Then

                        reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                        reconlineTable.Rows(i).Item("Accepted Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle Name").ToString.Trim)
                    End If
                    If UCase(vendor.item("First Name").ToString.Trim.Contains(surname)) Then
                        reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                        reconlineTable.Rows(i).Item("Accepted Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle Name").ToString.Trim)
                    End If
                    '   If (UCase(reconlineTable.Rows(i).Item("Surname")) = UCase(vendor.item("Surname").ToString)) And (UCase(reconlineTable.Rows(i).Item("First Name")) = UCase(vendor.item("First Name"))) And (UCase(vendor.item("Middle Names").ToString) = "") And (reconlineTable.Rows(i).Item("Middle Name") = "") Then
                    '  reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                    ' reconlineTable.Rows(i).Item("Accepted Name") = UCase(reconlineTable.Rows(i).Item("Surname") + " " + reconlineTable.Rows(i).Item("First Name") + " " + reconlineTable.Rows(i).Item("Middle Name"))
                    'End If


                    'WRONG PERIOD
                    Dim pin, per As String
                    pin = reconlineTable.Rows(i).Item("PIN")
                    per = reconlineTable.Rows(i).Item("period").ToString
                    If per = "" Then per = Today.ToString
                    '  lines = (From d In dbrecon.Recon_Engine_Lines Where d.PIN = pin And d.Period = per And d.No <> headerno Order By d.Period Descending).FirstOrDefault




                    Dim liness
                    liness = (From K In lineds.Tables(0).AsEnumerable
                              Where K.Field(Of String)("PIN").Contains(pin)).FirstOrDefault


                    If liness Is Nothing = False Then
                        If per = pery Then
                            reconlineTable.Rows(i).Item("Wrong Period") = True
                            reconlineTable.Rows(i).Item("Status") = 4  'wrong period
                        End If
                        'PREVIOUS AMOUNT
                        'type of remittance = monthly
                        If reconheader.Type_of_Remittance = 0 Then
                            reconlineTable.Rows(i).Item("PREVIOUS AMOUNT") = liness.Item("Total Contribution").ToString
                        End If
                    Else
                        reconlineTable.Rows(i).Item("PREVIOUS AMOUNT") = 0
                    End If
                End If

                'FOR BLOCKED PIN
                Try
                    If vendor.item("Blocked") = 2 Then
                        reconlineTable.Rows(i).Item("Blocked PIN") = 1
                        reconlineTable.Rows(i).Item("Status") = 5
                    End If
                Catch ex As Exception

                End Try


                Dim duplic = (From K In reconlineTable.AsEnumerable
                              Where K.Field(Of String)("PIN").Contains(reconlineTable.Rows(i).Item("PIN"))).Count
                If duplic > 1 Then
                    reconlineTable.Rows(i).Item("Status") = 3
                    reconlineTable.Rows(i).Item("Wrong Period") = False
                    reconlineTable.Rows(i).Item("Duplicate Record") = True
                    'duplicate
                End If


            Next
            reconlineTable.AcceptChanges()
            '  dbrecon.SubmitChanges()
            Dim queries As String = ""
            queries = " DELETE [dbo].[Recon Engine Lines Temp] WHERE [No] ='" & headerno & "'"
            returndataset("RECONDB", queries)

            ' db.SubmitChanges()

            Dim con1 As New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            con1.Open()

            Using bulkCopy As SqlBulkCopy =
              New SqlBulkCopy(con1)
                bulkCopy.DestinationTableName = "RECONDB.dbo.[Recon Engine Lines Temp]"

                Try
                    ' Write from the source to the destination.
                    bulkCopy.WriteToServer(reconlineTable)

                Catch ex As Exception
                    con1.Close()
                    Return (ex.Message)

                End Try
                ' con1.Close()
            End Using

            queries = ""
            queries += "    merge into [RECONDB].[dbo].[Recon Engine Lines] as Target"
            queries += " Using [RECONDB].[dbo].[Recon Engine Lines Temp] As Source"
            queries += " On Target.[No]=Source.[No]  And Target.[Line No]=Source.[Line No]"
            queries += " When matched Then "
            queries += " update set Target.[Suggested First Name]=Source.[Suggested First Name], Target.[Suggested Middle Name]= Source.[Suggested Middle Name], Target.[Suggested Surname] = Source.[Suggested Surname],"
            queries += " Target.[Status] = Source.[Status], Target.[PREVIOUS AMOUNT] = Source.[PREVIOUS AMOUNT], Target.[Wrong Period] = Source.[Wrong Period], Target.[Duplicate Record] = Source.[Duplicate Record],  Target.[Accepted Name] = Source.[Accepted Name],Target.[Blocked PIN] = Source.[Blocked PIN],"
            queries += " Target.[Not on NAV DB] = Source.[Not on NAV DB];"
            ' queries += " When Not matched Then"
            'queries += " insert([No]) values (Source.[No]);"
            ' returndataset("RECONDB", queries)

            Dim com As New SqlCommand
            com = New SqlCommand(queries, con1)
            If com.ExecuteNonQuery <> 0 Then
                con1.Close()
                queries = " DELETE [RECONDB].[dbo].[Recon Engine Lines Temp] WHERE [No] ='" & headerno & "'"
                returndataset("RECONDB", queries)

                con1.Dispose()
                reconheader.Status = 1
                reconheader.Reconciliation_DB = 1
                db.SubmitChanges()
            End If


            Return ""
        Catch ex As Exception
            Return "Error " & ex.Message
        End Try
    End Function
End Class
