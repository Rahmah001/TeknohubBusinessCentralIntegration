Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports Microsoft.SharePoint.Client

Imports System.Security
Imports System.Net
Imports System.IO


Public Class NAVSharePoint
    Public Sub New()
    End Sub
#Region "=====================================RunTimeParameters"
    'INSTANT VB NOTE: The variable sharePointSiteUrl was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private sharePointSiteUrl_Renamed As String
    'INSTANT VB NOTE: The variable documentLibrary was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private documentLibrary_Renamed As String
    'INSTANT VB NOTE: The variable userName was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private userName_Renamed As String
    'INSTANT VB NOTE: The variable password was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private password_Renamed As String
    'INSTANT VB NOTE: The variable domain was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private domain_Renamed As String
    'INSTANT VB NOTE: The variable filePathLog was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private filePathLog_Renamed As String
    'INSTANT VB NOTE: The variable documentPath was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private documentPath_Renamed As String
    Public Property SharePointSiteUrl() As String
        Get
            Return sharePointSiteUrl_Renamed
        End Get
        Set(ByVal value As String)
            sharePointSiteUrl_Renamed = value
        End Set
    End Property
    Public Property DocumentLibrary() As String
        Get
            Return documentLibrary_Renamed
        End Get
        Set(ByVal value As String)
            documentLibrary_Renamed = value
        End Set
    End Property
    Public Property UserName() As String
        Get
            Return userName_Renamed
        End Get
        Set(ByVal value As String)
            userName_Renamed = value
        End Set
    End Property
    Public Property Password() As String
        Get
            Return password_Renamed
        End Get
        Set(ByVal value As String)
            password_Renamed = value
        End Set
    End Property
    Public Property Domain() As String
        Get
            Return domain_Renamed
        End Get
        Set(ByVal value As String)
            domain_Renamed = value
        End Set
    End Property
    Public Property DocumentPath() As String
        Get
            Return documentPath_Renamed
        End Get
        Set(ByVal value As String)
            documentPath_Renamed = value
        End Set
    End Property
    Public Property FilePathLog() As String
        Get
            Return filePathLog_Renamed
        End Get
        Set(ByVal value As String)
            filePathLog_Renamed = value
        End Set
    End Property
#End Region ' ====================RRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRRR
#Region "--------------------------------------Metadata Parameters"
    'INSTANT VB NOTE: The variable rsaPin was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private rsaPin_Renamed As String
    'INSTANT VB NOTE: The variable firstName was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private firstName_Renamed As String
    'INSTANT VB NOTE: The variable surname was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private surname_Renamed As String
    'INSTANT VB NOTE: The variable otherNames was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private otherNames_Renamed As String
    'INSTANT VB NOTE: The variable employerName was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private employerName_Renamed As String
    'INSTANT VB NOTE: The variable mobileNo was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private mobileNo_Renamed As String
    'INSTANT VB NOTE: The variable nextofKin was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private nextofKin_Renamed As String
    'INSTANT VB NOTE: The variable employerCode was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private employerCode_Renamed As String
    'INSTANT VB NOTE: The variable agentCode was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private agentCode_Renamed As String
    'INSTANT VB NOTE: The variable documentType was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private documentType_Renamed As String
    Public Property DocumentType() As String
        Get
            Return documentType_Renamed
        End Get
        Set(ByVal value As String)
            documentType_Renamed = value
        End Set
    End Property
    Public Property RSAPin() As String
        Get
            Return rsaPin_Renamed
        End Get
        Set(ByVal value As String)
            rsaPin_Renamed = value
        End Set
    End Property
    Public Property FirstName() As String
        Get
            Return firstName_Renamed
        End Get
        Set(ByVal value As String)
            firstName_Renamed = value
        End Set
    End Property
    Public Property Surname() As String
        Get
            Return surname_Renamed
        End Get
        Set(ByVal value As String)
            surname_Renamed = value
        End Set
    End Property
    Public Property OtherNames() As String
        Get
            Return otherNames_Renamed
        End Get
        Set(ByVal value As String)
            otherNames_Renamed = value
        End Set
    End Property
    Public Property EmployerName() As String
        Get
            Return employerName_Renamed
        End Get
        Set(ByVal value As String)
            employerName_Renamed = value
        End Set
    End Property
    Public Property MobileNo() As String
        Get
            Return mobileNo_Renamed
        End Get
        Set(ByVal value As String)
            mobileNo_Renamed = value
        End Set
    End Property
    Public Property NextofKin() As String
        Get
            Return nextofKin_Renamed
        End Get
        Set(ByVal value As String)
            nextofKin_Renamed = value
        End Set
    End Property
    Public Property EmployerCode() As String
        Get
            Return employerCode_Renamed
        End Get
        Set(ByVal value As String)
            employerCode_Renamed = value
        End Set
    End Property
    Public Property AgentCode() As String
        Get
            Return agentCode_Renamed
        End Get
        Set(ByVal value As String)
            agentCode_Renamed = value
        End Set
    End Property
#End Region ' ++++++++++++++++++++++++++++++++++++++++++++MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM

    Public Function uploadToSharePoint() As Response
        SPUtilityHandle.FilePathLOG = Me.FilePathLog
        Dim thisresponse As New Response()
        Try



            Dim siteUrl As String = Me.SharePointSiteUrl
            Dim sourcePath As String = Me.DocumentPath
            'Insert Credentials
            Dim context As New ClientContext(siteUrl)
            Dim passWord As New SecureString()
            For Each c In Me.Password
                passWord.AppendChar(c)
            Next c
            context.Credentials = New NetworkCredential(Me.UserName, passWord, Me.Domain)

            Dim site As Microsoft.SharePoint.Client.Web = context.Web
            'Get the required RootFolder
            Dim barRootFolderRelativeUrl As String = Me.DocumentLibrary
            Dim currentRunFolder As Folder
            Try
                currentRunFolder = site.GetFolderByServerRelativeUrl(barRootFolderRelativeUrl)
                Try
                    Dim newFile As FileCreationInformation = New FileCreationInformation With {
                        .Content = System.IO.File.ReadAllBytes(sourcePath),
                        .Url = Path.GetFileName(sourcePath),
                        .Overwrite = True
                    }
                    currentRunFolder.Files.Add(newFile)
                    currentRunFolder.Update()
                    Try
                        context.ExecuteQuery()
                        thisresponse.Indicator = False
                        thisresponse.ResponseMessage = "File Successfully uploaded"
                    Catch ex As Exception
                        SPUtilityHandle.Logger("SPLOG", "File upload Query failure " & ex.Message)
                        thisresponse.Indicator = False
                        thisresponse.ResponseMessage = "File upload Query failure " & ex.Message
                    End Try
                    'Folder currentRunFile = site.GetFolderByServerRelativeUrl(barRootFolderRelativeUrl + "/" + newFolderName + "/" + Path.GetFileName(@p));
                    Dim myurl As String = siteUrl & "/" & barRootFolderRelativeUrl & "/" & Path.GetFileName(sourcePath)
                    Dim doclist As Microsoft.SharePoint.Client.List = site.Lists.GetByTitle(Me.DocumentLibrary)
                    Dim item As Microsoft.SharePoint.Client.ListItem = SPUtilityHandle.LoadItembyUrl(doclist, myurl)
                    Try
                        item("RSAPin") = Me.RSAPin
                        item("FirstName") = Me.FirstName
                        item("Surname") = Me.Surname
                        item("OtherNames") = Me.OtherNames
                        item("EmployerName") = Me.EmployerName
                        item("MobileNo") = Me.MobileNo
                        item("NextofKin") = Me.NextofKin
                        item("EmployerCode") = Me.EmployerCode
                        item("AgentCode") = Me.AgentCode
                        item("DocumentType") = Me.DocumentType
                        item.Update()
                    Catch ex As Exception
                        SPUtilityHandle.Logger("SPLOG", "Item search failure " & ex.Message)
                        thisresponse.Indicator = False
                        thisresponse.ResponseMessage = "Item search failure  " & ex.Message
                    End Try
                    Try
                        context.ExecuteQuery()
                        thisresponse.Indicator = True
                        thisresponse.ResponseMessage = "File Successfully uploaded with metadata"
                        thisresponse.Url = siteUrl & "/" & barRootFolderRelativeUrl & "/" & Path.GetFileName(sourcePath)
                    Catch ex As Exception
                        SPUtilityHandle.Logger("SPLOG", "Metadata update failure " & ex.Message)
                        thisresponse.Indicator = False
                        thisresponse.ResponseMessage = "Metadata update failure " & ex.Message
                    End Try
                Catch ex As Exception
                    Try


                        SPUtilityHandle.Logger("SPLOG", "Either File or SharePoint Library not found " & ex.Message)
                    Catch ex1 As Exception

                    End Try
                    thisresponse.Indicator = False
                    thisresponse.ResponseMessage = "Either File or SharePoint Library not Found"
                End Try
            Catch exec As Exception
                SPUtilityHandle.Logger("SPLOG", "Folder not found " & exec.Message)
                thisresponse.Indicator = False
                thisresponse.ResponseMessage = "Folder not found " & Me.DocumentLibrary
            End Try

            Return thisresponse
        Catch ex As Exception
            SPUtilityHandle.Logger("SPLOG", "Metadata update failure " & ex.Message)
            thisresponse.Indicator = False
            thisresponse.ResponseMessage = "Metadata update failure " & ex.Message
        End Try
        Return thisresponse
    End Function
End Class

