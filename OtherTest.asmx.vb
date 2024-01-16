Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.DirectoryServices


' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class OtherTest
    Inherits System.Web.Services.WebService

    <WebMethod()> _
    Public Function HelloWorld() As String
       Return "Hello World"
    End Function
    Dim Dsearch As DirectorySearcher
    <WebMethod>
    Public Function getADMails() As String
        'Get the username and domain information
        Dim userName As String = Environment.UserName
        Dim domainName As String = Environment.UserDomainName

        'Set the correct format for the AD query and filter
        'Dim ldapQueryFormat As String = "LDAP://{0}.com/DC={0},DC=com"
        Dim ldapQueryFormat As String = "LDAP://192.168.0.12"
        Dim rootQuery As String = String.Format(ldapQueryFormat, "armpension")
      

        Dim root As New DirectoryEntry(rootQuery)
        Dsearch = New DirectorySearcher(root)
     
        getarray(root)

        Return getalldetails2("armpension\navsupport")
    End Function
    Public Shared Function GetProperty(ByVal searchResult As SearchResult, ByVal PropertyName As String) As String
       
        If searchResult.Properties.Contains(PropertyName) Then
            Return PropertyName & " : " & searchResult.Properties(PropertyName)(0).ToString()
        Else
            Return String.Empty
        End If


       
    End Function
    Dim availablereq As New ArrayList
    Public Function getalldetails(ByVal username As String) As String
        getalldetails = ""
        ' Dim mName As String = username
      '   Dim queryFilterFormat As String = "(&(samAccountName={0})(objectCategory=person)(objectClass=user))"
        '    Dim searchFilter As String = String.Format(queryFilterFormat, "Rahmot")

        ' Dsearch.Filter = "(&(objectClass=user)(l=" + username + "))"



        ' get all entries from the active directory.
        ' Last Name, name, initial, homepostaladdress, title, company etc..
        For Each sResultSet As SearchResult In Dsearch.FindAll()
            ' Display Name
            getalldetails += (GetProperty(sResultSet, "DisplayName")) & vbCrLf
            ' userPrincipalName
            getalldetails += (GetProperty(sResultSet, "userPrincipalName")) & vbCrLf

            ' First Name
            getalldetails += (GetProperty(sResultSet, "samAccountName")) & vbCrLf
            ' Login Name
            getalldetails += (GetProperty(sResultSet, "cn")) & vbCrLf
            ' First Name
            getalldetails += (GetProperty(sResultSet, "givenName")) & vbCrLf
            ' Middle Initials
            getalldetails += (GetProperty(sResultSet, "initials")) & vbCrLf
            ' Last Name
            getalldetails += (GetProperty(sResultSet, "sn")) & vbCrLf
            ' Address
            Dim tempAddress As String = GetProperty(sResultSet, "homePostalAddress")

            If tempAddress <> String.Empty Then
                Dim addressArray() As String = tempAddress.Split(";"c)
                Dim taddr1, taddr2 As String
                taddr1 = addressArray(0)
                getalldetails += (taddr1)
                taddr2 = addressArray(1)
                getalldetails += (taddr2)
            End If
            ' title
            getalldetails += (GetProperty(sResultSet, "title")) & vbCrLf
            ' company
            getalldetails += (GetProperty(sResultSet, "company")) & vbCrLf
            'state
            getalldetails += (GetProperty(sResultSet, "st")) & vbCrLf
            'city
            getalldetails += (GetProperty(sResultSet, "l")) & vbCrLf
            'country
            getalldetails += (GetProperty(sResultSet, "co")) & vbCrLf
            'postal code
            getalldetails += (GetProperty(sResultSet, "postalCode")) & vbCrLf
            ' telephonenumber
            getalldetails += (GetProperty(sResultSet, "telephoneNumber")) & vbCrLf
            'extention
            getalldetails += (GetProperty(sResultSet, "otherTelephone")) & vbCrLf
            'fax
            getalldetails += (GetProperty(sResultSet, "facsimileTelephoneNumber")) & vbCrLf

            ' email address
            getalldetails += (GetProperty(sResultSet, "mail")) & vbCrLf
            ' Challenge Question
            getalldetails += (GetProperty(sResultSet, "extensionAttribute1")) & vbCrLf
            ' Challenge Response
            getalldetails += (GetProperty(sResultSet, "extensionAttribute2")) & vbCrLf
            'Member Company
            getalldetails += (GetProperty(sResultSet, "extensionAttribute3")) & vbCrLf
            ' Company Relation ship Exits
            getalldetails += (GetProperty(sResultSet, "extensionAttribute4")) & vbCrLf
            'status
            getalldetails += (GetProperty(sResultSet, "extensionAttribute5")) & vbCrLf
            ' Assigned Sales Person
            getalldetails += (GetProperty(sResultSet, "extensionAttribute6")) & vbCrLf
            ' Accept T and C
            getalldetails += (GetProperty(sResultSet, "extensionAttribute7")) & vbCrLf
            ' jobs
            getalldetails += (GetProperty(sResultSet, "extensionAttribute8")) & vbCrLf
            Dim tEamil As String = GetProperty(sResultSet, "extensionAttribute9")

            ' email over night
            If tEamil <> String.Empty Then
                Dim em1, em2, em3 As String
                Dim emailArray() As String = tEamil.Split(";"c)
                em1 = emailArray(0)
                em2 = emailArray(1)
                em3 = emailArray(2)
                getalldetails += (em1 & em2 & em3)

            End If
            ' email daily emerging market
            getalldetails += (GetProperty(sResultSet, "extensionAttribute10")) & vbCrLf
            ' email daily corporate market
            getalldetails += (GetProperty(sResultSet, "extensionAttribute11")) & vbCrLf
            ' AssetMgt Range
            getalldetails += (GetProperty(sResultSet, "extensionAttribute12")) & vbCrLf
            ' date of account created
            getalldetails += (GetProperty(sResultSet, "whenCreated")) & vbCrLf
            ' date of account changed
            getalldetails += (GetProperty(sResultSet, "whenChanged")) & vbCrLf
        Next sResultSet
    End Function
    Public Function getalldetails2(ByVal username As String) As String
        getalldetails2 = ""
       
        For Each sResultSet As SearchResult In Dsearch.FindAll()
            ' Display Name
            For i As Integer = 0 To availablereq.Count - 1
                getalldetails2 += GetProperty(sResultSet, availablereq(i)) & "   "
            Next

        Next sResultSet
    End Function

    Public Shared Function PrintDirectoryEntryProperties(entry As System.DirectoryServices.DirectoryEntry) As String
        PrintDirectoryEntryProperties = 0
        ' loop through all the properties and get the key for each
        For Each Key As String In entry.Properties.PropertyNames
            Dim sPropertyValues As String = [String].Empty
            ' now loop through all the values in the property;
            ' can be a multi-value property
            For Each Value As Object In entry.Properties(Key)
                sPropertyValues += Convert.ToString(Value) & ";"
            Next
            ' cut off the separator at the end of the value list
            sPropertyValues = sPropertyValues.Substring(0, sPropertyValues.Length - 1)
            ' now add the property info to the property list
            PrintDirectoryEntryProperties += (Key & "=" & sPropertyValues)
        Next
    End Function
    Public Sub getarray(ByVal de As DirectoryEntry)

        For Each strAttrName As String In de.Properties.PropertyNames
            availablereq.Add(strAttrName)
        Next
       
    End Sub
     
End Class