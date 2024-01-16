Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports XLeratorDLL_financial
Imports System.Data.SqlClient


' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class InvestService
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function
    <WebMethod()>
    Public Function GetResult(ByVal Code As String) As Double
        ' GetResult = 10
        Dim cfAmt As Double()
        Dim cfDate As DateTime()
        Dim result As Double
        Dim cont As Integer = 0
        Dim cs As String = My.MySettings.Default.ARM_TESTINGConnectionString


        Using con As SqlConnection = New SqlConnection(cs)

            Using cmd As SqlCommand = New SqlCommand()
                cmd.Connection = con
                cmd.CommandType = CommandType.Text
                cmd.CommandText = "select date,amt from [AIICO FUND ACCOUNTS$cs table] where [Investment code]='" & Code & "'"
                con.Open()
                Dim da As SqlDataAdapter = New SqlDataAdapter(cmd)
                Dim dt As DataTable = New DataTable()
                da.Fill(dt)
                cfAmt = New Double(dt.Rows.Count - 1) {}
                cfDate = New DateTime(dt.Rows.Count - 1) {}

                For Each row As DataRow In dt.Rows
                    cfDate(cont) = row.Field(Of DateTime)(0)
                    cfAmt(cont) = Convert.ToDouble(row.Field(Of Decimal)(1))
                    cont += 1
                Next


                result = XLeratorDLL_financial.XLeratorDLL_financial.XIRR(cfAmt, cfDate)
                con.Close()
                Console.WriteLine(result)
                Return result
            End Using
        End Using
    End Function
End Class