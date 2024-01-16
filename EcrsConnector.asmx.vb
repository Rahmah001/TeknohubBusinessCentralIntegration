Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Drawing
Imports System.Drawing.Drawing2D

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class EcrsConnector
    Inherits System.Web.Services.WebService
    <WebMethod()>
    Public Function getNinConsent(ByVal temporaryid As String, ByVal niNconsentpdf As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.convertNIN(temporaryid, niNconsentpdf)
        Return resp
    End Function
    <WebMethod()>
    Public Function getNinConsentNew(ByVal temporaryid As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.getNINJpeg(temporaryid)
        Return resp
    End Function
    <WebMethod()>
    Public Function getNinConsentCFINew(ByVal temporaryid As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.getNINJpegCFI(temporaryid)
        Return resp
    End Function
    <WebMethod()>
    Public Function getNinConsentcfi(ByVal temporaryid As String, ByVal niNconsentpdf As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.convertNINRecapture(temporaryid, niNconsentpdf)
        Return resp
    End Function

    <WebMethod()>
    Public Function SaveExistingNinConsent(ByVal temporaryid As String, ByVal niNconsentjpg As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.SaveExistingNin(temporaryid, niNconsentjpg)
        Return resp
    End Function
    <WebMethod()>
    Public Function SaveExistingNinConsentCFI(ByVal temporaryid As String, ByVal niNconsentjpg As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        Dim CONVP As New Pdf2Jpg
        resp = CONVP.SaveExistingNinRecapure(temporaryid, niNconsentjpg)
        Return resp
    End Function
    <WebMethod()>
    Public Function GetPINfromECRS(ByVal temporaryid As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        resp = penin.GeneratePIN(temporaryid)
        Return resp
    End Function
    <WebMethod()>
    Public Function SendRecapturetoECRS(ByVal pin As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        resp = penin.PencomRecapture(pin)
        Return resp

    End Function

    <WebMethod()>
    Public Function GetUpdatewithSetid(ByVal setid As String, ByVal uniqueno As String, ByVal requesttype As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        resp = penin.getPencomRequestStatus(setid, uniqueno, requesttype)
        Return resp
    End Function
    <WebMethod()>
    Public Function sendRecordUpdate(ByVal ucode As String, ByVal pin As String, ByVal SN As String, ByVal FN As String,
                                     ByVal Fieldcodes As String, ByVal oldvals As String, ByVal newvals As String) As String
        Dim penin As New PencomIntegrator
        Dim resp As String
        resp = penin.PencomUpdateBiodata(ucode, pin, FN, SN, Fieldcodes, oldvals, newvals)
        Return resp

    End Function

End Class