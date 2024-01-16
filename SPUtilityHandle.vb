Option Infer On

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Microsoft.SharePoint.Client
Imports System.IO


Public Module SPUtilityHandle
    'INSTANT VB NOTE: The variable filePathLog was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private filePathLog_Renamed As String
    Public Property FilePathLOG() As String
        Get
            Return filePathLog_Renamed
        End Get
        Set(ByVal value As String)
            filePathLog_Renamed = value
        End Set
    End Property
    <System.Runtime.CompilerServices.Extension> _
    Public Function LoadItembyUrl(ByVal list As List, ByVal url As String) As ListItem
        Dim context = list.Context
        Dim query = New CamlQuery With {.ViewXml = String.Format("<View><RowLimit>1</RowLimit><Query><Where><Eq><FieldRef Name='FileRef'/><Value Type='Url'>{0}</Value></Eq></Where></Query></View>", url)}
        Dim items = list.GetItems(query)
        context.Load(items)
        Try
            context.ExecuteQuery()
        Catch ex As Exception
            Logger("SPLOG", ex.Message)
        End Try
        Return If(items.Count > 0, items(0), Nothing)
    End Function
    Public Sub Logger(ByVal FileName As String, ByVal ErrorMessage As String)
        Dim LogFile As String = FilePathLOG

        ' dd/mm/yyyy hh:mm:ss AM/PM ==> Log Message
        Dim sLogFormat As String = Date.Now.ToShortDateString().ToString() & " " & Date.Now.ToLongTimeString().ToString() & " ==> "

        'ErrorLogYYYYMMDD
        Dim sYear As String = Date.Now.Year.ToString()
        Dim sMonth As String = Date.Now.Month.ToString()
        Dim sDay As String = Date.Now.Day.ToString()
        Dim sErrorTime As String = sYear & sMonth & sDay

        Try
            Dim sw As New StreamWriter(LogFile & FileName & sErrorTime, True)
            sw.WriteLine(sLogFormat & ErrorMessage)
            sw.Flush()
            sw.Close()
        Catch Ex As Exception
            Logger("GENERAL SCHEMA LOG", " Log File Failure " & Ex.Message)
        End Try
    End Sub
End Module
