Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.DirectoryServices
Imports System.DirectoryServices.ActiveDirectory
Imports System.DirectoryServices.AccountManagement

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class GeneralServices
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld() As String
        Return "Hello World"
    End Function
    <WebMethod()> _
    Public Function returnEmail(ByVal username As String) As String
        Dim uName As String = ""
        uName = accounttest(username.Replace("armpension\", ""))
        Return uName
    End Function

    Private Function uEmail(ByVal uid As String) As String
        Dim dirSearcher As New DirectorySearcher()
        ' Dim entry As New DirectoryEntry(dirSearcher.SearchRoot.Path)
        Dim entry As New DirectoryEntry("LDAP://192.168.0.12")


        'Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
        'Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)
        'Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel



        dirSearcher.Filter = "(&(objectClass=user)(objectcategory=person)(username=" & uid & "*))"

        Dim srEmail As SearchResult = dirSearcher.FindOne()

        Dim propName As String = "mail"
        Dim valColl As ResultPropertyValueCollection = srEmail.Properties(propName)
        Try
            Return valColl(0).ToString()
        Catch
            Return ""
        End Try

    End Function
    Private Function ValidateActiveDirectoryLogin(ByVal Domain As String, ByVal Username As String, ByVal Password As String) As Boolean
        Dim Success As Boolean = False
        Dim Entry As New System.DirectoryServices.DirectoryEntry("LDAP://" & Domain, Username, Password)
        Dim Searcher As New System.DirectoryServices.DirectorySearcher(Entry)
        Searcher.SearchScope = DirectoryServices.SearchScope.OneLevel
        Try
            Dim Results As System.DirectoryServices.SearchResult = Searcher.FindOne
            Success = Not (Results Is Nothing)
        Catch ex As Exception
            MsgBox(ex.Message)
            Success = False
        End Try
        Return Success
    End Function
    Public Function accounttest(ByVal username As String) As String
        Using de As New DirectoryEntry("LDAP://192.168.0.12")
            Using adSearch As New DirectorySearcher(de)
                adSearch.Filter = "(sAMAccountName=" & username & ")"
                Dim adSearchResult As SearchResult = adSearch.FindOne()
                Dim propName As String = "mail"
                Dim valColl As ResultPropertyValueCollection = adSearchResult.Properties(propName)
                Try
                    Return valColl(0).ToString()
                Catch
                    Return ""
                End Try
            End Using
        End Using
      
    End Function
    Public Function AuthenticateUser(ByVal uname As String, ByVal pass As String) As Boolean
        Dim username As String = uname
        Dim password As String = pass
        'Dim domain As String = 'this can be in a config file, hard coded (I wouldnt do that), or inputed from the UI
        Dim domain As String = "10.234.240.140"

        Dim isAuthenticated As Boolean = ValidateActiveDirectoryLogin(domain, username, password)

        Return isAuthenticated
    End Function


End Class