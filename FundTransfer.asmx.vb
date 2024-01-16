Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Data.SqlClient
Imports System.Reflection

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class FundTransfer
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function
    <WebMethod()>
    Public Function ReconEngines(ByVal headerno As String) As String
        Try


            Dim recEngineLinestring As String = ""

            Dim db As New DataClasses1DataContext
            Dim db1 As New accountingdDataContext
            Dim dbrecon As New reconlinesclassDataContext

            Dim vendorquery As String
            vendorquery = "SELECT No_, P_I_N, Surname,[Fund ID], 
                        [Date of Birth], FLOOR((CAST (GetDate() AS INTEGER) - CAST([Date of Birth] AS INTEGER)) / 365.25) AS Age from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] 
                        where FLOOR((CAST (GetDate() AS INTEGER) - CAST([Date of Birth] AS INTEGER)) / 365.25) < 50 and P_I_N like 'PEN%' and LEN(P_I_N)=15  and [Fund ID] = 1"

            Dim dsvendor As New DataSet
            dsvendor = returndataset("ARM", vendorquery)




            Dim reconheader = (From i In db.AIICO_FUND_ACCOUNTS_Reconciliation_Engines
                               Where i.No = headerno).FirstOrDefault
            If reconheader Is Nothing Then
                Return "Error Connecting to recon engine db"
            End If

            Dim reconLines = (From i In dbrecon.Recon_Engine_Lines Where i.No = headerno)
            Dim reconlineTable As New DataTable
            reconlineTable = ToDataTables(reconLines)

            If reconLines Is Nothing Then
                Return "No line records"
            End If



            Dim linecount As Integer = reconLines.Count


            Dim pinstring As String = ""




            Dim query = "select No_ ,P_I_N ,Surname ,[First Name] ,[Middle Names] from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] where [P_I_N] in ("
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
            '   Dim vendorlist = (From k In db1.ARM_FUND_ACCOUNTS_Vendors Where list.Contains(k.P_I_N) Select k.No_, k.First_Name, k.Surname, k.Middle Names, k.P_I_N)

            ' For Each item As Recon_Engine_Line In reconLines
            For i As Integer = 0 To reconlineTable.Rows.Count - 1
                vendor = (From K In vendords.Tables(0).AsEnumerable
                          Where K.Field(Of String)("P_I_N").Contains(reconlineTable.Rows(i).Item("PIN"))).FirstOrDefault


                '   If reconlineTable.Rows(i).Item("Line_No") = 144119 Then
                '  MsgBox("here")
                ' End If

                If vendor Is Nothing Then

                    reconlineTable.Rows(i).Item("Not_on_NAV_DB") = True
                    reconlineTable.Rows(i).Item("Status") = 2 ' not in db
                    reconlineTable.Rows(i).Item("Accepted_Name") = ""
                    reconlineTable.Rows(i).Item("Suggested_Surname") = ""
                    reconlineTable.Rows(i).Item("Suggested_First_Name") = ""
                    reconlineTable.Rows(i).Item("Suggested_Middle_Name") = ""
                    reconlineTable.Rows(i).Item("Apply_Suggestions") = False
                    ' Return "" '"PIN DOESNT EXIST IN THE DB"
                Else
                    reconlineTable.Rows(i).Item("Status") = 0
                    reconlineTable.Rows(i).Item("Accepted_Name") = ""
                    reconlineTable.Rows(i).Item("Suggested_Surname") = ""
                    reconlineTable.Rows(i).Item("Suggested_First_Name") = ""
                    reconlineTable.Rows(i).Item("Suggested_Middle_Name") = ""
                    reconlineTable.Rows(i).Item("Apply_Suggestions") = False

                    'Dim sur, sur1 As String
                    If UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim) <> UCase(vendor.item("Surname").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested_Surname") = vendor.item("Surname").ToString.Trim
                        '  reconlinetable.rows(i).item("Suggested_Surname = vendor.item("Surname").ToString()
                        '  sur = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim)
                        ' sur1 = UCase(vendor.item("Surname").ToString.Trim)

                    End If
                    If UCase(reconlineTable.Rows(i).Item("First_Name").ToString.Trim) <> UCase(vendor.item("First Name").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested_First_Name") = vendor.item("First Name").ToString.Trim
                    End If
                    If UCase(reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim) <> UCase(vendor.item("Middle Names").ToString.Trim) Then
                        reconlineTable.Rows(i).Item("Suggested_Middle_Name") = vendor.item("Middle Names").ToString()
                    End If




                    If (UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim) = UCase(vendor.item("Surname").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("First_Name").ToString.Trim) = UCase(vendor.item("First Name").ToString.Trim)) And (UCase(reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim) = UCase(vendor.item("Middle Names").ToString.Trim)) Then

                        reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                        reconlineTable.Rows(i).Item("Accepted_Name") = UCase(reconlineTable.Rows(i).Item("Surname").ToString.Trim + " " + reconlineTable.Rows(i).Item("First_Name").ToString.Trim + " " + reconlineTable.Rows(i).Item("Middle_Name").ToString.Trim)
                    End If

                    '   If (UCase(reconlineTable.Rows(i).Item("Surname")) = UCase(vendor.item("Surname").ToString)) And (UCase(reconlineTable.Rows(i).Item("First_Name")) = UCase(vendor.item("First Name"))) And (UCase(vendor.item("Middle Names").ToString) = "") And (reconlineTable.Rows(i).Item("Middle_Name") = "") Then
                    '  reconlineTable.Rows(i).Item("Status") = 1 'MATCHED
                    ' reconlineTable.Rows(i).Item("Accepted_Name") = UCase(reconlineTable.Rows(i).Item("Surname") + " " + reconlineTable.Rows(i).Item("First_Name") + " " + reconlineTable.Rows(i).Item("Middle_Name"))
                    'End If


                    'WRONG PERIOD
                    Dim pin, per As String
                    pin = reconlineTable.Rows(i).Item("PIN")
                    per = reconlineTable.Rows(i).Item("period")

                    '  lines = (From d In dbrecon.Recon_Engine_Lines Where d.PIN = pin And d.Period = per And d.No <> headerno Order By d.Period Descending).FirstOrDefault




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

                Dim duplic = (From K In reconlineTable.AsEnumerable
                              Where K.Field(Of String)("PIN").Contains(reconlineTable.Rows(i).Item("PIN"))).Count
                If duplic > 1 Then
                    reconlineTable.Rows(i).Item("Status") = 3
                    reconlineTable.Rows(i).Item("Wrong_Period") = False
                    reconlineTable.Rows(i).Item("Duplicate_Record") = True
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
            con1.Open

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
            queries += " Target.[Status] = Source.[Status], Target.[PREVIOUS AMOUNT] = Source.[PREVIOUS AMOUNT], Target.[Wrong Period] = Source.[Wrong Period], Target.[Duplicate Record] = Source.[Duplicate Record],  Target.[Accepted Name] = Source.[Accepted Name],"
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

    Private Function returndataset(ByVal databasenames As String, ByVal query As String) As DataSet
        returndataset = Nothing
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
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


            Return Nothing
        End Try
    End Function
End Class