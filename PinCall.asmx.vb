Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports Microsoft.SharePoint.Client

' To allow this Web Service to be called from script, using ASP.NET AJAX, uncomment the following line.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://attain-es.com/")>
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)>
<ToolboxItem(False)>
Public Class PinCall
    Inherits System.Web.Services.WebService

    <WebMethod()>
    Public Function sendsms(ByVal url As String) As String
        Try
            ' Create a request using a URL that can receive a post. 

            Dim request As WebRequest = WebRequest.Create(url)
            ' Set the Method property of the request to POST.

            request.Method = "POST"

            ' Create POST data and convert it to a byte array.
            Dim obj As Object
            obj = Split(url, "?")

            Dim postData As String = obj(1).ToString

            Dim byteArray As Byte() = Encoding.UTF8.GetBytes(postData)
            ' Set the ContentType property of the WebRequest.
            ' request.ContentType = "application/x-www-form-urlencoded"
            request.ContentType = "application/json;charset=UTF-8"
            request.Proxy = Nothing

            ' Set the ContentLength property of the WebRequest.
            request.ContentLength = byteArray.Length
            ' Get the request stream.
            Dim dataStream As Stream = request.GetRequestStream()
            ' Write the data to the request stream.
            dataStream.Write(byteArray, 0, byteArray.Length)
            ' Close the Stream object.
            dataStream.Close()
            ' Get the response.

            Dim response As WebResponse = request.GetResponse()

            ' Display the status.
            '    Console.WriteLine(CType(response, HttpWebResponse).StatusDescription)
            ' Get the stream containing content returned by the server.
            dataStream = response.GetResponseStream()
            ' Open the stream using a StreamReader for easy access.
            Dim reader As New StreamReader(dataStream)
            ' Read the content.
            Dim responseFromServer As String = reader.ReadToEnd()

            ' Clean up the streams.
            reader.Close()
            dataStream.Close()
            response.Close()
            Return CType(response, HttpWebResponse).StatusDescription
        Catch ex As Exception
            Return "Error: " & ex.Message
        End Try

    End Function
    <WebMethod()>
    Public Function sendMail(updateType As String, pin As String, ByVal emailto As String, ByVal emailcc As String, ByVal username As String, ByVal requesttime As String, ByVal status As String) As String

        Dim url As String = "http://41.216.173.228/armotherservice/activationPage.aspx"

        Dim msg As String '"Dear " + fullname + " ,"

        msg = "An Update Request have been " & status & " for PIN:" + pin + " by " + username + " on " + updateType & " at " + requesttime
        msg += "<br/><br/><br/><br/>"
        Dim message As New MailMessage()
        Try


            message.[To].Add(emailto)
            message.CC.Add(emailcc)
            message.Bcc.Add("abdulmalik.musa@armpension.com")
            message.From = New MailAddress("workflow@armpension.com", "ARM Pensions WorkFlow")

            message.Subject = "ARM Pension Update Request for PIN:" & pin
            message.Body = msg
            message.IsBodyHtml = True
            Dim client As New SmtpClient()

            client.Credentials = New System.Net.NetworkCredential(My.MySettings.Default.Unames, My.MySettings.Default.Pnames)
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Port = 25
            client.Host = My.MySettings.Default.Smmm

            client.Send(message)
        Catch ex As Exception
            Return ex.Message.ToString()
            Return False
        End Try
        Return ""
    End Function
    <WebMethod()>
    Public Function sendMailwithreason(updateType As String, pin As String, ByVal emailto As String, ByVal emailcc As String, ByVal username As String, ByVal requesttime As String, ByVal status As String, ByVal comment As String) As String


        Dim msg As String '"Dear " + fullname + " ,"

        msg = "An Update Request have been " & status & " for PIN:" + pin + " by " + username + " on " + updateType & " at " + requesttime & " because " + comment
        msg += "<br/><br/><br/><br/>"
        Dim message As New MailMessage()
        Try

            message.[To].Add(emailto)
            message.CC.Add(emailcc)
            message.Bcc.Add("abdulmalik.musa@armpension.com")
            message.From = New MailAddress("workflow@armpension.com", "ARM Pensions WorkFlow")

            message.Subject = "ARM Pension Update Request for PIN:" & pin
            message.Body = msg
            message.IsBodyHtml = True
            Dim client As New SmtpClient()

            client.Credentials = New System.Net.NetworkCredential(My.MySettings.Default.Unames, My.MySettings.Default.Pnames)
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Port = 25
            client.Host = My.MySettings.Default.Smmm

            client.Send(message)
        Catch ex As Exception
            Return ex.Message.ToString()
            Return False
        End Try
        Return ""
    End Function
    <WebMethod()>
    Public Function sendEscalationMail(updateType As String, pin As String, ByVal email As String, ByVal username As String, ByVal requesttime As String, ByVal status As String) As String

        Dim url As String = "http://41.216.173.228/armotherservice/activationPage.aspx"

        Dim msg As String '"Dear " + fullname + " ,"

        msg = "<b> List of requests that have exceeded the time limit.</b>"
        msg += "<br/><br/><br/><br/>"
        Dim resp As String = getEmailTablestring("Modulename", updateType, pin, email, username, requesttime, status)
        msg += resp
        Dim message As New MailMessage()
        Try

            message.Bcc.Add("abdulmalik.musa@armpension.com")
            message.Bcc.Add("rahmot.oshoala@gmail.com")

            message.[To].Add(email)
            message.From = New MailAddress("workflow@armpension.com", "ARM Pensions WorkFlow")

            message.Subject = "ARM Pension Escalation mail on Update Request for  PIN:" & pin
            message.Body = msg
            message.IsBodyHtml = True
            Dim client As New SmtpClient()

            client.Credentials = New System.Net.NetworkCredential(My.MySettings.Default.Unames, My.MySettings.Default.Pnames)
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Port = 25


            client.Host = My.MySettings.Default.Smmm


            client.Send(message)
        Catch ex As Exception
            Return "Error:" & ex.Message.ToString()
            Return False
        End Try
        Return ""
    End Function
    <WebMethod()>
    Public Function sendOtherEscalationMail(ByVal ModuleName As String, updateType As String, pin As String, ByVal email As String, ByVal username As String, ByVal requesttime As String, ByVal status As String, ByVal Subj As String) As String

        ' Dim url As String = "http://41.216.173.228/armotherservice/activationPage.aspx"
        '   Dim url As String = "http://192.168.0.30/armotherservice/activationPage.aspx"

        Dim msg As String '"Dear " + fullname + " ,"

        msg = "<b> List of requests that have exceeded the time limit.</b>"
        msg += "<br/><br/><br/><br/>"
        Dim resp As String = getEmailTablestring(ModuleName, updateType, pin, email, username, requesttime, status)
        msg += resp
        Dim message As New MailMessage()
        Try



            message.[To].Add(email)
            ' message.CC.Add(emailcc)
            message.Bcc.Add("abdulmalik.musa@armpension.com")
            message.From = New MailAddress("workflow@armpension.com", "ARM Pensions WorkFlow")

            message.Subject = Subj '"ARM Pension Escalation mail on Update Request for  PIN:" & pin
            message.Body = msg
            message.IsBodyHtml = True
            Dim client As New SmtpClient()

            client.Credentials = New System.Net.NetworkCredential(My.MySettings.Default.Unames, My.MySettings.Default.Pnames)
            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Port = 25
            client.Host = My.MySettings.Default.Smmm

            client.Send(message)
        Catch ex As Exception
            Return ex.Message.ToString()
            Return False
        End Try
        Return ""
    End Function
    Private Function getEmailTablestring(Modulename As String, updateType As String, pin As String, ByVal email As String, ByVal username As String, ByVal requesttime As String, ByVal status As String) As String


        Dim rows As Integer = Split(updateType, "|").Count - 1
        Dim cols As Integer = 6

        Dim strHeader As String = "<table  width='600px' align='center' border='1' cellpadding='1' cellspacing='1' style='border-top:5px solid white;><tbody style='font-size:12px;font-family:Trebuchet MS;'>"
        Dim strFooter As String = "</tbody></table>"
        Dim sbContent As New StringBuilder()
        Dim heading1 As String
        If Modulename = "BENEFIT" Then heading1 = "REQUEST TYPE" Else heading1 = "UPDATE TYPE"

        '------------header rows
        sbContent.Append("<tr>")
        sbContent.Append(String.Format("<td>{0}</td>", heading1))
        sbContent.Append(String.Format("<td>{0}</td>", "PIN"))
        sbContent.Append(String.Format("<td>{0}</td>", "USER ID"))
        sbContent.Append(String.Format("<td>{0}</td>", "REQUEST DATE TIME"))
        sbContent.Append(String.Format("<td>{0}</td>", "CURRENT STATUS"))
        sbContent.Append("</tr>")

        For i As Integer = 0 To rows
            sbContent.Append("<tr>")
            sbContent.Append(String.Format("<td>{0}</td>", Split(updateType, "|")(i)))
            sbContent.Append(String.Format("<td>{0}</td>", Split(pin, "|")(i)))
            sbContent.Append(String.Format("<td>{0}</td>", Split(username, "|")(i)))
            sbContent.Append(String.Format("<td>{0}</td>", Split(requesttime, "|")(i)))
            sbContent.Append(String.Format("<td>{0}</td>", Split(status, "|")(i)))
            sbContent.Append("</tr>")
        Next i

        Dim emailTemplate As String = strHeader & sbContent.ToString() & strFooter
        Return emailTemplate
    End Function
    <WebMethod()>
    Public Function sendsms2Infobip(ByVal mobileno As String, ByVal smsbody As String, ByVal sendersrole As String) As String
        Dim SS As New Infobip
        Dim resp As String = SS.sendsmsInfobip(mobileno, smsbody)
        Return resp
    End Function




    '<WebMethod()>
    'Public Function UploadDoctoSharepoint(ByVal MylongString As String) As String
    '    Dim SharePointCons = New NewSharepointRequestClass()
    '    Dim respons As New NewSharepointRespClass
    '    Dim context As ClientContext
    '    Dim objvariable As Object = Split(MylongString, "|")
    '    Try
    '        SharePointCons.SharePointSiteUrl = objvariable(0)
    '        SharePointCons.UserName = objvariable(1)
    '        SharePointCons.Domain = objvariable(2)
    '        SharePointCons.Password = objvariable(3)
    '        SharePointCons.DocumentLibrary = objvariable(4)
    '        SharePointCons.DocumentPath = objvariable(5)
    '        SharePointCons.DocumentType = objvariable(6)
    '        SharePointCons.FilePathLog = objvariable(7)
    '        SharePointCons.FirstName = objvariable(8)
    '        SharePointCons.OtherNames = objvariable(9)
    '        SharePointCons.Surname = objvariable(10)
    '        SharePointCons.RSAPin = objvariable(11)
    '        SharePointCons.MobileNo = objvariable(12)
    '        SharePointCons.EmployerCode = objvariable(13)
    '        SharePointCons.EmployerName = objvariable(14)
    '        SharePointCons.NextofKin = objvariable(15)


    '        context = New ClientContext(SharePointCons.SharePointSiteUrl)
    '        ' Dim list As Microsoft.SharePoint.Client.List = context.Web.Lists.GetByTitle("Nav-Documents")
    '        Dim list As Microsoft.SharePoint.Client.List = context.Web.Lists.GetByTitle(SharePointCons.DocumentLibrary)

    '        Dim cred As NetworkCredential = New NetworkCredential()
    '        cred.UserName = SharePointCons.UserName
    '        cred.Password = SharePointCons.Password
    '        cred.Domain = SharePointCons.Domain
    '        context.Credentials = cred
    '        Dim fileCI As FileCreationInformation = New FileCreationInformation()
    '        Dim direct As String = SharePointCons.DocumentPath
    '        Dim fs As FileStream = New FileStream(direct, FileMode.Open, FileAccess.Read)
    '        Dim br As BinaryReader = New BinaryReader(fs)
    '        Dim filecontents As Byte() = br.ReadBytes(CInt(fs.Length))
    '        br.Close()
    '        fs.Close()
    '        fileCI.Content = filecontents
    '        fileCI.Overwrite = True
    '        fileCI.Url = System.IO.Path.GetFileName(direct)
    '        list.RootFolder.Files.Add(fileCI)
    '        context.ExecuteQuery()
    '        respons.indicator = True
    '    Catch ex As Exception
    '        respons.indicator = False
    '        respons.ResponseMessage = ex.Message
    '        Return respons.indicator & "|" & respons.ResponseMessage & "|" & respons.Url
    '    End Try
    '    Dim myurl As String = SharePointCons.SharePointSiteUrl & "/" & SharePointCons.DocumentLibrary & "/" & Path.GetFileName(SharePointCons.DocumentPath)
    '    Dim doclist As Microsoft.SharePoint.Client.List = context.Web.Lists.GetByTitle(SharePointCons.DocumentLibrary)
    '    Dim item As Microsoft.SharePoint.Client.ListItem = SPUtilityHandle.LoadItembyUrl(doclist, myurl)


    '    Try
    '        item("RSAPin") = SharePointCons.RSAPin
    '        item("FirstName") = SharePointCons.FirstName
    '        item("Surname") = SharePointCons.Surname
    '        item("OtherNames") = SharePointCons.OtherNames
    '        item("EmployerName") = SharePointCons.EmployerName
    '        item("MobileNo") = SharePointCons.MobileNo
    '        item("NextofKin") = SharePointCons.NextofKin
    '        item("EmployerCode") = SharePointCons.EmployerCode
    '        item("AgentCode") = SharePointCons.AgentCode
    '        item("DocumentType") = SharePointCons.DocumentType
    '        item.Update()
    '    Catch ex As Exception
    '        respons.indicator = False
    '        respons.ResponseMessage = "Item search failure  " & ex.Message
    '        Return respons.indicator & "|" & respons.ResponseMessage & "|" & respons.Url
    '    End Try

    '    Try
    '        context.ExecuteQuery()
    '        respons.indicator = True
    '        respons.ResponseMessage = "File Successfully uploaded with metadata"
    '        respons.Url = myurl
    '    Catch ex As Exception
    '        respons.indicator = False
    '        respons.ResponseMessage = "Metadata update failure " & ex.Message
    '        Return respons.indicator & "|" & respons.ResponseMessage & "|" & respons.Url
    '    End Try
    '    Return respons.indicator & "|" & respons.ResponseMessage & "|" & respons.Url
    'End Function
    '<WebMethod()>
    'Public Function moveRecordtoSharePoint(ByVal MylongString As String) As String
    '    Try
    '        Dim SharePointCons = New NAVSharePoint()
    '        moveRecordtoSharePoint = Nothing
    '        Dim objvariable As Object = Split(MylongString, "|")

    '        SharePointCons.SharePointSiteUrl = objvariable(0)
    '        SharePointCons.UserName = objvariable(1)
    '        SharePointCons.Domain = objvariable(2)
    '        SharePointCons.Password = objvariable(3)
    '        SharePointCons.DocumentLibrary = objvariable(4)
    '        SharePointCons.DocumentPath = objvariable(5)
    '        SharePointCons.DocumentType = objvariable(6)
    '        SharePointCons.FilePathLog = objvariable(7)
    '        SharePointCons.FirstName = objvariable(8)
    '        SharePointCons.OtherNames = objvariable(9)
    '        SharePointCons.Surname = objvariable(10)
    '        SharePointCons.RSAPin = objvariable(11)
    '        SharePointCons.MobileNo = objvariable(12)
    '        SharePointCons.EmployerCode = objvariable(13)
    '        SharePointCons.EmployerName = objvariable(14)
    '        SharePointCons.NextofKin = objvariable(15)

    '        Dim responsecons As Response = SharePointCons.uploadToSharePoint()
    '        moveRecordtoSharePoint = responsecons.Indicator & "|" &
    '        responsecons.ResponseMessage & "|" & responsecons.Url

    '    Catch ex As Exception
    '        moveRecordtoSharePoint = ex.Message
    '    End Try
    '    Return moveRecordtoSharePoint
    'End Function
    <WebMethod()>
    Public Function updateMobileTable(ByVal easyconstring As String, ByVal code As String) As String
        Dim query As String
        Dim dst As New DataSet
        Try


            query = "Select * from [Mobile_Table] where code = '" & code & "'"
            dst = returndataset(easyconstring, query)

            If dst.Tables(0).Rows.Count = 0 Then

                Return Nothing
            End If

            Dim smscon As New SqlClient.SqlConnection
            smscon.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            smscon.Open()
            smscon.ChangeDatabase(easyconstring)
            Dim cmdcom As New SqlClient.SqlCommand
            Dim RIGHT_TOMB As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RIGHT_TOMB")))
            Dim LEFT_TOMB As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("LEFT_TOMB")))
            Dim Picture As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("Picture")))
            Dim SIGNATURE As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("SIGNATURE")))
            Dim RIGHT_INDEX As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RIGHT_INDEX")))
            Dim LEFT_INDEX As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("LEFT_INDEX")))
            Dim RIGHT_MIDDLE As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RIGHT_MIDDLE")))
            Dim LEFT_MIDDLE As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("LEFT_MIDDLE")))
            Dim RIGHT_RING As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RIGHT_RING")))
            Dim LEFT_RING As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("LEFT_RING")))
            Dim RIGHT_LITTLE As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RIGHT_LITTLE")))
            Dim LEFT_LITTLE As Byte() = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("LEFT_LITTLE")))
            Dim PID As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("PID")) = False Then PID = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("PID")))
            Dim POA As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("POA")) = False Then POA = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("POA")))
            Dim Digital_Signature As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("Digital_Signature")) = False Then Digital_Signature = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("Digital_Signature")))
            Dim RSA_Statement As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("RSA_Statement")) = False Then RSA_Statement = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("RSA_Statement")))
            Dim Birth_Certificate As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("Birth_Certificate")) = False Then Birth_Certificate = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("Birth_Certificate")))
            Dim Medical_Report As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("Medical_Report")) = False Then Medical_Report = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("Medical_Report")))
            Dim Full_Picture As Byte() = Nothing
            If IsDBNull(dst.Tables(0).Rows(0).Item("Full_Picture")) = False Then Full_Picture = System.Convert.FromBase64String(vstring(dst.Tables(0).Rows(0).Item("Full_Picture")))

            '        query = "insert into Mobile_Biometrics (CODE, RIGHT_TOMB, LEFT_TOMB, Picture, Signature, RIGHT_INDEX, LEFT_INDEX, RIGHT_MIDDLE, LEFT_MIDDLE, RIGHT_RING, LEFT_RING, RIGHT_LITTLE, LEFT_LITTLE, Digital_Signature, PID, POA, RSA_Statement, Birth_Certificate, Medical_Report, Full_Picture) 
            'values ('" + code + "',@RIGHT_TOMB,@LEFT_TOMB,@Picture,@SIGNATURE,@RIGHT_INDEX,@LEFT_INDEX,@RIGHT_MIDDLE,@LEFT_MIDDLE,@RIGHT_RING,@LEFT_RING,@RIGHT_LITTLE,@LEFT_LITTLE,@Digital_Signature,@PID,@POA,@RSA_Statement,@Birth_Certificate,@Medical_Report,@Full_Picture)"

            query = "update Mobile_Biometrics set RIGHT_TOMB= @RIGHT_TOMB,LEFT_TOMB=@LEFT_TOMB,Picture=@Picture,
SIGNATURE=@SIGNATURE,RIGHT_INDEX=@RIGHT_INDEX,LEFT_INDEX=@LEFT_INDEX,RIGHT_MIDDLE=@RIGHT_MIDDLE,LEFT_MIDDLE=@LEFT_MIDDLE,
RIGHT_RING=@RIGHT_RING,LEFT_RING=@LEFT_RING,RIGHT_LITTLE=@RIGHT_LITTLE,LEFT_LITTLE=@LEFT_LITTLE"
            If IsNothing(POA) = False Then query += ",POA =@POA"
            If IsNothing(PID) = False Then query += ",PID =@PID"
            If IsNothing(Digital_Signature) = False Then query += ",Digital_Signature =@Digital_Signature"
            If IsNothing(RSA_Statement) = False Then query += " ,RSA_Statement=@RSA_Statement"
            If IsNothing(Birth_Certificate) = False Then query += "  ,Birth_Certificate=@Birth_Certificate"
            If IsNothing(Medical_Report) = False Then query += " ,Medical_Report=@Medical_Report"
            If IsNothing(Full_Picture) = False Then query += " ,Full_Picture=@Full_Picture"
            query += " where [code] = '" & code & "'"

            cmdcom = New SqlCommand(query, smscon)
            cmdcom.Parameters.Add("@RIGHT_TOMB", SqlDbType.Image).Value = RIGHT_TOMB
            cmdcom.Parameters.Add("@LEFT_TOMB", SqlDbType.Image).Value = LEFT_TOMB
            cmdcom.Parameters.Add("@Picture", SqlDbType.Image).Value = Picture
            cmdcom.Parameters.Add("@SIGNATURE", SqlDbType.Image).Value = SIGNATURE
            cmdcom.Parameters.Add("@RIGHT_INDEX", SqlDbType.Image).Value = RIGHT_INDEX
            cmdcom.Parameters.Add("@LEFT_INDEX", SqlDbType.Image).Value = LEFT_INDEX
            cmdcom.Parameters.Add("@RIGHT_MIDDLE", SqlDbType.Image).Value = RIGHT_MIDDLE
            cmdcom.Parameters.Add("@LEFT_MIDDLE", SqlDbType.Image).Value = LEFT_MIDDLE
            cmdcom.Parameters.Add("@RIGHT_RING", SqlDbType.Image).Value = RIGHT_RING
            cmdcom.Parameters.Add("@LEFT_RING", SqlDbType.Image).Value = LEFT_RING
            cmdcom.Parameters.Add("@RIGHT_LITTLE", SqlDbType.Image).Value = RIGHT_LITTLE
            cmdcom.Parameters.Add("@LEFT_LITTLE", SqlDbType.Image).Value = LEFT_LITTLE
            If IsNothing(Digital_Signature) = False Then cmdcom.Parameters.Add("@Digital_Signature", SqlDbType.Image).Value = Digital_Signature
            If IsNothing(PID) = False Then cmdcom.Parameters.Add("@PID", SqlDbType.Image).Value = PID
            If IsNothing(POA) = False Then cmdcom.Parameters.Add("@POA", SqlDbType.Image).Value = POA
            If IsNothing(RSA_Statement) = False Then cmdcom.Parameters.Add("@RSA_Statement", SqlDbType.Image).Value = RSA_Statement
            If IsNothing(Birth_Certificate) = False Then cmdcom.Parameters.Add("@Birth_Certificate", SqlDbType.Image).Value = Birth_Certificate
            If IsNothing(Medical_Report) = False Then cmdcom.Parameters.Add("@Medical_Report", SqlDbType.Image).Value = Medical_Report
            If IsNothing(Full_Picture) = False Then cmdcom.Parameters.Add("@Full_Picture", SqlDbType.Image).Value = Full_Picture

            
            If cmdcom.ExecuteNonQuery() > 0 Then
                smscon.Close()
                smscon.Dispose()
                Return "Success"
            End If
            smscon.Close()
            smscon.Dispose()
        Catch ex As Exception
            Return ex.Message

        End Try
        Return "nothing"
    End Function
    Public Function vstring(ByVal values As String) As String
        Dim begining As String = Mid(values, 1, 22)
        If begining.EndsWith(",") Then
            Return Mid(values, 23)
        Else
            Return Mid(values, 24)
        End If

    End Function
    '<WebMethod()>
    'Public Function pdfManage(ByVal sourcePDFpath As String) As String
    '    'Dim WorkingFolder As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
    '    'Dim InputFile As String = Path.Combine(WorkingFolder, "Test.pdf")
    '    'Dim OutputFile As String = Path.Combine(WorkingFolder, "Test_enc.pdf")
    '    Dim directory As String = Path.GetDirectoryName(sourcePDFpath)
    '    Dim filename As String = Path.GetFileNameWithoutExtension(sourcePDFpath)
    '    Dim InputFile As String = sourcePDFpath

    '    Dim OutputFile As String = directory + "\" + filename + "T.pdf"
    '    Dim pinstring As String = Mid(filename, 1, 15)
    '    Dim partpinstring As String = Mid(pinstring, Len(pinstring) - 3, 4)

    '    Dim input As Stream = New FileStream(InputFile, FileMode.Open, FileAccess.Read, FileShare.Read)
    '    Dim output As Stream = New FileStream(OutputFile, FileMode.Create, FileAccess.Write, FileShare.None)
    '    Dim reader As PdfReader = New PdfReader(input)
    '    PdfEncryptor.Encrypt(reader, output, True, partpinstring, partpinstring, PdfWriter.ALLOW_SCREENREADERS)
    '    Return OutputFile
    'End Function
    '<WebMethod()>
    Public Function MoveBiometrics(ByVal easyconstring As String, ByVal pin As String) As String
        Dim query As String
        Dim dst As New DataSet
        Dim tempid As String
        Dim vendorno As String
        Dim enrolleeds As New DataSet

        Try

            query = "Select No_,P_I_N,[Form Reference Number] from [Temporary Client]  WITH (NOLOCK) where P_I_N = '" & pin & "' select No_ from  [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor]  WITH (NOLOCK) where  P_I_N = '" & pin & "' "

            dst = returndataset(easyconstring, query)

            If dst.Tables.Count <= 1 Then
                Return "Error: Connection Issue"
            End If
            If dst.Tables(0).Rows.Count = 0 Then

                Return "Error: Connection Issue on Temporary Client"
            End If

            If dst.Tables(1).Rows.Count = 0 Then
                Return "Error: Connection Issue on Vendor Table"
            End If
            If dst.Tables(0).Rows.Count = 0 Or dst.Tables(1).Rows.Count = 0 Then
                Return "Error: Not moved yet"
            End If
            tempid = dst.Tables(0).Rows(0).Item("No_")
            vendorno = dst.Tables(1).Rows(0).Item(0)
            Try

                enrolleeds = returndataset(easyconstring, "select * from enrollee where code ='" & dst.Tables(0).Rows(0).Item(2).ToString & "'")
            Catch ex As Exception

            End Try

            query = "Select * from [empImg]  WITH (NOLOCK) where PIN_N = '" & tempid & "' or REGISTRATION_CODE_N ='" & vendorno & "'"
            dst = returndataset(easyconstring, query)


            Dim smscon As New SqlClient.SqlConnection
            smscon.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            smscon.Open()
            smscon.ChangeDatabase(easyconstring)
            Dim cmdcom As New SqlClient.SqlCommand



            If dst.Tables(0).Rows.Count = 0 Then

                query = "INSERT INTO dbo.empimg (REGISTRATION_CODE,REGISTRATION_CODE_N,PIN,PIN_N,PICTURE_IMAGE,SIGNATURE_IMAGE,THUMBPRINT,THUMBPRINT1 )
Select  '" & vendorno & "','" & vendorno & "','" & pin & "','" & tempid & "',PICTURE,SIGNATURE,RIGHT_TOMB,LEFT_TOMB From dbo.[Temporary Biometrics] Where TEMPORARY_ID ='" & tempid & "'"

                dst = returndataset(easyconstring, query)


                query = "INSERT INTO [dbo].[Complete_Bio]([CODE], [Vendor_No], [RIGHT_TOMB], [LEFT_TOMB], [PICTURE], [Signature], [RIGHT_INDEX], [LEFT_INDEX], [RIGHT_MIDDLE], [LEFT_MIDDLE], [RIGHT_RING], [LEFT_RING], [RIGHT_LITTLE], [LEFT_LITTLE], [ENROLLE_FK], [Digital_Signature], [PID], [POA], EXPATRAITE_DOCUMENT, FINGERPRINT_DOCUMENT, EMPLOYMENT_CONFIRMATION_DOC,
LETTER_APPOINTMENT,TRANSFER_ACCEPTANCE_SERVICE,BIRTH_CERTIFICATE,PROMOTION_LETTER_SLIP,PROMOTION_LETTER_SLIP_04,PROMOTION_LETTER_SLIP_07,
PROMOTION_LETTER_SLIP_10,PROMOTION_LETTER_SLIP_13,PROMOTION_LETTER_SLIP_16,LETTER_FIRST_APPOINTMENT,LETTER_APPOINTMENT_PRIVATE,Full_Picture,RSA_STATEMENT,Scanned_RSA_Form,Scanned_RSA_Form_2)
     Select  '" & tempid & "','" & vendorno & "',[RIGHT_TOMB],[LEFT_TOMB],[PICTURE],[SIGNATURE],[RIGHT_INDEX],[LEFT_INDEX],[RIGHT_MIDDLE],[LEFT_MIDDLE],[RIGHT_RING],[LEFT_RING],[RIGHT_LITTLE],[LEFT_LITTLE],[ENROLLE_FK],[Digital_Signature],[PID],[POA],EXPATRAITE_DOCUMENT,FINGERPRINT_DOCUMENT,EMPLOYMENT_CONFIRMATION_DOC,
LETTER_APPOINTMENT,TRANSFER_ACCEPTANCE_SERVICE,BIRTH_CERTIFICATE,PROMOTION_LETTER_SLIP,PROMOTION_LETTER_SLIP_04,PROMOTION_LETTER_SLIP_07,
PROMOTION_LETTER_SLIP_10,PROMOTION_LETTER_SLIP_13,PROMOTION_LETTER_SLIP_16,LETTER_FIRST_APPOINTMENT,LETTER_APPOINTMENT_PRIVATE,Full_Picture,RSA_STATEMENT,Scanned_RSA_Form,Scanned_RSA_Form_2 From dbo.[Temporary Biometrics] Where TEMPORARY_ID ='" & tempid & "'"

                dst = returndataset(easyconstring, query)

            Else
                query = "Select * from [Complete_Bio]  WITH (NOLOCK) where CODE = '" & tempid & "'"
                Dim dscomplete = returndataset(easyconstring, query)

                If dscomplete.Tables(0).Rows.Count = 0 Then
                    query = "INSERT INTO [dbo].[Complete_Bio]([CODE], [Vendor_No], [RIGHT_TOMB], [LEFT_TOMB], [PICTURE], [Signature], [RIGHT_INDEX], [LEFT_INDEX], [RIGHT_MIDDLE], [LEFT_MIDDLE], [RIGHT_RING], [LEFT_RING], [RIGHT_LITTLE], [LEFT_LITTLE], [ENROLLE_FK], [Digital_Signature], [PID], [POA], EXPATRAITE_DOCUMENT, FINGERPRINT_DOCUMENT, EMPLOYMENT_CONFIRMATION_DOC,
LETTER_APPOINTMENT,TRANSFER_ACCEPTANCE_SERVICE,BIRTH_CERTIFICATE,PROMOTION_LETTER_SLIP,PROMOTION_LETTER_SLIP_04,PROMOTION_LETTER_SLIP_07,
PROMOTION_LETTER_SLIP_10,PROMOTION_LETTER_SLIP_13,PROMOTION_LETTER_SLIP_16,LETTER_FIRST_APPOINTMENT,LETTER_APPOINTMENT_PRIVATE,full_picture,RSA_STATEMENT,Scanned_RSA_Form,Scanned_RSA_Form_2)
     Select  '" & tempid & "','" & vendorno & "',[RIGHT_TOMB],[LEFT_TOMB],[PICTURE],[SIGNATURE],[RIGHT_INDEX],[LEFT_INDEX],[RIGHT_MIDDLE],[LEFT_MIDDLE],[RIGHT_RING],[LEFT_RING],[RIGHT_LITTLE],[LEFT_LITTLE],[ENROLLE_FK],[Digital_Signature],[PID],[POA],EXPATRAITE_DOCUMENT,FINGERPRINT_DOCUMENT,EMPLOYMENT_CONFIRMATION_DOC, LETTER_APPOINTMENT,TRANSFER_ACCEPTANCE_SERVICE,BIRTH_CERTIFICATE,PROMOTION_LETTER_SLIP,PROMOTION_LETTER_SLIP_04,PROMOTION_LETTER_SLIP_07,
PROMOTION_LETTER_SLIP_10,PROMOTION_LETTER_SLIP_13,PROMOTION_LETTER_SLIP_16,LETTER_FIRST_APPOINTMENT,LETTER_APPOINTMENT_PRIVATE,Full_Picture,RSA_STATEMENT,Scanned_RSA_Form,Scanned_RSA_Form From dbo.[Temporary Biometrics] Where TEMPORARY_ID ='" & tempid & "'"

                    dst = returndataset(easyconstring, query)

                End If

                query = " UPDATE empimg  SET  REGISTRATION_CODE = '" & vendorno & "',REGISTRATION_CODE_N = '" & vendorno & "', PIN= '" & pin & "', PIN_N='" & tempid & "',PICTURE_IMAGE=PICTURE,SIGNATURE_IMAGE=Signature,THUMBPRINT= RIGHT_TOMB,THUMBPRINT1=LEFT_TOMB  FROm empimg  INNER Join  [Temporary Biometrics] On empimg.pin_n =[Temporary Biometrics].temporary_id
                WHERE   TEMPORARY_ID ='" & tempid & "' or REGISTRATION_CODE_N='" & vendorno & "'"
                dst = returndataset(easyconstring, query)

                query = "UPDATE Complete_Bio  Set [RIGHT_TOMB]=tb.RIGHT_TOMB, [LEFT_TOMB]=tb.LEFT_TOMB, [PICTURE]=tb.PICTURE, [Signature]=tb.Signature, [RIGHT_INDEX]=tb.RIGHT_INDEX, [LEFT_INDEX]=tb.LEFT_INDEX, [RIGHT_MIDDLE]=tb.RIGHT_MIDDLE, [LEFT_MIDDLE]=tb.LEFT_MIDDLE, 
[RIGHT_RING] = tb.RIGHT_RING, [LEFT_RING]=tb.LEFT_RING, [RIGHT_LITTLE]=tb.RIGHT_LITTLE, [LEFT_LITTLE]=tb.LEFT_LITTLE, [ENROLLE_FK]=tb.ENROLLE_FK, [Digital_Signature]=tb.Digital_Signature, [PID]=tb.PID, [POA]=tb.POA,
EXPATRAITE_DOCUMENT = tb.EXPATRAITE_DOCUMENT , FINGERPRINT_DOCUMENT=tb.FINGERPRINT_DOCUMENT , EMPLOYMENT_CONFIRMATION_DOC=tb.EMPLOYMENT_CONFIRMATION_DOC,
LETTER_APPOINTMENT=tb.LETTER_APPOINTMENT,TRANSFER_ACCEPTANCE_SERVICE=tb.TRANSFER_ACCEPTANCE_SERVICE,BIRTH_CERTIFICATE=tb.BIRTH_CERTIFICATE,PROMOTION_LETTER_SLIP=tb.PROMOTION_LETTER_SLIP,PROMOTION_LETTER_SLIP_04=tb.PROMOTION_LETTER_SLIP_04,
PROMOTION_LETTER_SLIP_07=tb.PROMOTION_LETTER_SLIP_07,PROMOTION_LETTER_SLIP_10=tb.PROMOTION_LETTER_SLIP_10,PROMOTION_LETTER_SLIP_13=tb.PROMOTION_LETTER_SLIP_13,PROMOTION_LETTER_SLIP_16=tb.PROMOTION_LETTER_SLIP_16,LETTER_FIRST_APPOINTMENT=tb.LETTER_FIRST_APPOINTMENT,LETTER_APPOINTMENT_PRIVATE=tb.LETTER_APPOINTMENT_PRIVATE,
FULL_PICTURE=tb.FULL_PICTURE,RSA_STATEMENT=tb.RSA_STATEMENT,Scanned_RSA_Form=tb.Scanned_RSA_Form,Scanned_RSA_Form_2=tb.Scanned_RSA_Form_2
from Complete_Bio  INNER Join  [Temporary Biometrics]tb  On Complete_Bio.code =tb.temporary_id
              WHERE   TEMPORARY_ID ='" & tempid & "'"
                dst = returndataset(easyconstring, query)



            End If


            If enrolleeds Is Nothing Then Return "success"
            If enrolleeds.Tables.Count = 0 Then Return "success"
            If enrolleeds.Tables(0).Rows.Count = 0 Then Return "success"

            query = "UPDATE Complete_Bio  Set WSQ_LEFT_INDEX  ='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_LEFT_INDEX") & "', WSQ_LEFT_LITTLE='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_LEFT_LITTLE") & "', WSQ_LEFT_MIDDLE='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_LEFT_MIDDLE") & "',
WSQ_LEFT_RING='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_LEFT_RING") & "', WSQ_LEFT_TOMB='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_LEFT_TOMB") & "', 
WSQ_RIGHT_INDEX='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_RIGHT_INDEX") & "', WSQ_RIGHT_LITTLE='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_RIGHT_LITTLE") & "',
WSQ_RIGHT_MIDDLE='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_RIGHT_MIDDLE") & "', 
WSQ_RIGHT_RING = '" & enrolleeds.Tables(0).Rows(0).Item("WSQ_RIGHT_RING") & "', WSQ_RIGHT_TOMB='" & enrolleeds.Tables(0).Rows(0).Item("WSQ_RIGHT_TOMB") & "'  WHERE   CODE ='" & tempid & "' "
            dst = returndataset(easyconstring, query)
        Catch ex As Exception
            Return ex.Message

        End Try
        Return "success"
    End Function
    <WebMethod>
    Function MovetoVendor(ByVal no As String, ByVal pin As String) As String
        Dim query, querynok, queryben, queryvenadd As String
        Dim ds As New DataSet
        ds = returndataset("easyregserver", "select * from " & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] where P_I_N = '" & pin & "'")
        If IsNothing(ds) = False Then
            If ds.Tables(0).Rows.Count <> 0 Then
                Return "Error: Record Moved"
            End If
        End If

        Dim vendornods As DataSet = returndataset("", "select CONVERT(VARCHAR(50),CAST([Last No_ Used] AS decimal) + 1 ) from [" & My.MySettings.Default.COMPANYNAME.Trim & "$No_ Series Line] where [Series Code] ='MEMBERSHIP'")
        Dim vendorno As String = vendornods.Tables(0).Rows(0).Item(0)

        query = "
        insert into [" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor] 
([No_],	[Name],	[Search Name],	[Name 2],	[Address],	[Address 2],	[City],	[Contact],	[Phone No_],	[Telex No_],	
[Our Account No_],	[Territory Code],	[Global Dimension 1 Code],	[Global Dimension 2 Code],	[Budgeted Amount],	[Vendor Posting Group],	[Currency Code],	[Language Code],	[Statistics Group],	[Payment Terms Code],
	[Fin_ Charge Terms Code],	[Purchaser Code],	[Shipment Method Code],	[Shipping Agent Code],	[Invoice Disc_ Code],	[Country_Region Code],	[Blocked],	[Pay-to Vendor No_],	[Priority],	[Payment Method Code],
	[Last Date Modified],	[Application Method],	
	[Prices Including VAT],	[Fax No_],	[Telex Answer Back],	[VAT Registration No_],	[Gen_ Bus_ Posting Group],	[Picture],	[Post Code],	[County],	[E-Mail],	[Home Page],
		[No_ Series],	[Tax Area Code],	[Tax Liable],	[VAT Bus_ Posting Group],	[Block Payment Tolerance],	[IC Partner Code],	[Prepayment _],	[Cash Flow Payment Terms Code],	[Primary Contact No_],	[Responsibility Center],	
[Location Code],	[Lead Time Calculation],	[Base Calendar Code],	[Sex],	[Date of Birth],	[Pensionable Service Start Date],	[Date of Normal Retirement],	[DateMonth],	[Qualification],	[Vendor Type],	
		[Pension Type],	[Pensionable Service],	[Fund ID],	[Reason for Retirement],	[Retirement Date],	[ID_ No_],	[Resident Indicator],	[P_I_N],	[Marital Status],	[Membership Status],	
		[Year Interest Amount],	[NSSF Number],	[AVC Form Signed?],	[AVC _ of Salary],	[Scheme Join Date],	[Special Contribution],	[Date of First Employment],	[Payment Status],	[Agent Code],	[State of Origin],	[Local Government Authority],
		[Armed Forces],	[BirthDateMonth],	[Temporary PIN],	[Employer No_],	[Client Type],	[Branch Code],	[ID Type],	[ID Number],	[ID Expiry Date],	[Email Notification],	[Post Notification],	[Office Notification],	[SMS Notification],	[Bank Name],	[Annual Basic Allowance],	[Ann_ Transport Allowance],	[Ann_ Housing Allowance],	[Ann_ Other Allowance],	[EE Contribution],	[ER Contribution],	[VC Contribution],	[Created By],	[Approved By],	[Temp],	[Employer Name in Full],	[Staff File No_],	[Designation],	[State of Posting],	[Employer RC Number],	[Total Contribution],	[Vital Info],	[Form Reference Number],	[Created Date],	[Created By User ID],	[Last Modified Date],	[Last Modified By User ID],	[Religion],	[Surname],	[First Name],	[Middle Names],	[Title],	[Mobile 1],	[Office Phone No_],	[Grade Level],	[Step],	[Correspondence State Code],	[Statement Option],	[Nationality],	[Status],	[Verified],	[Risk Profile],	[Client Status],	[Occupation_Employment Status],	[Relationship Manager],	[Bank Sort Code],	[Bank State],	[Flag Code],	[Customer Income Frequency],	[Block_House No],	[Premises_Estate],	[Date in Address],	[ID Issue Date],	[Place of Issue],	[Bank Address],	[Bank Telephone No],	
		[Last Receipt Date],	[Dividend Mandate],	[RSA Number],	[Employer City],	[E_Mail 2],	[Vendor Type Finance],	[Employer Status],	[KBA Branch Code],	[Withholding Tax Code],	[Residential Address],	[State],	[Other Title],	[Correspondence Address],	[Correspondence Town],	[Correspondence Country],	[Upload Date],	[Date of Confirmation],	[Date of Last Promotion],	[Total Pensionable Salary],	[Salary Structure],	[Debtor Code],	[Payroll Status],	[Pays Tax?],	[Salary Scale],	[Correspondence Town2],	
[Pay Mode],	[Basic Pay],	[Employer Business Name],	[Industry Code],	[Employee Monthly Contribution],	[Employer Monthly Contribution],	[Total Monthly Contribution],	[Total Mon_ Pensionable Salary],	[Account Type],	[Employee Status],	[Employer State Code],	[Checked Date],	[Investproduct1],	[SchemeId],	[Invest percentage],	[Employer Address1],	[PIN Batch],	[Old PFA Code],	[LGA Name],	[State Name],	[State Of Posting Name],	[State Of Origin Name],	[Gender],	[Client type Asset],	[Payment Frequency],	[Select],	[PFA Code],	[Block RSA Statement],	[Block Retiree Statement],	[Military],	[Selected By],	[AVC Agent Code],	[Document Library No],	[KYC],	[Creditor No_]
,GLN,[Partner Type],[Image],[Preferred Bank Account Code],[Document Sending Profile])
Select '" & vendorno & "',	[Name],	'',	'',	[Address],	[Address 2],	[City],	'',	[Phone No_],	'',
    '',	'',	'',	'',	0,	'',	'',	'',	0,	'',	
            '',		'',	'',	'',	'',	[Country_Region Code],	[Blocked],	'',	0,	'',	
            '17530101',		[Application Method],	
0,	'',	'',	'',	'',	0x0,	'',	[County],	[E-Mail],	'',
	'',	'',	0,	'',	0,	'',	0,	'',	'',	'',	
	'','',	'',	[Sex],	[Date of Birth],	'17530101',	[Date of Normal Retirement],	[DateMonth],	[Qualification],	0,
		0,	0,	'1',	'',	'17530101',	'',	'',	[P_I_N],	[Marital Status],	[Membership Status],
		0,	0,	0,	0,	'17530101',	0,	[Date of First Employment],	[Payment Status],	[Agent Code],	[State of Origin],	[Local Government Authority],	
		0,	'',	'',	[Employer No_],	[Client Type],	[Branch Code],	[ID Type],	[ID Number],	[ID Expiry Date],	'',	'',	[Office Notification],	[SMS Notification],	'',	[Annual Basic Allowance],	[Ann_ Transport Allowance],	[Ann_ Housing Allowance],	[Ann_ Other Allowance],	[EE Contribution],	[ER Contribution],	[VC Contribution],	'',	[Approved By],	0,	[Employer Name in Full],	[Staff File No_],	[Designation],	[State of Posting],	[Employer RC Number],	[Total Contribution],	[Vital Info],	[Form Reference Number],	convert(date,[Created Date]),	[Created By User ID],	convert(date,[Last Modified Date]),	[Last Modified By User ID],	[Religion],	[Surname],	[First Name],	[Middle Names],	[Title],	[Mobile 1],	[Office Phone No_],	[Grade Level],	[Step],	[Correspondence State Code],	[Statement Option],	[Nationality],	[Status],	[Verified],	'',	'',	[Occupation_Employment Status],	[Relationship Manager],	[Bank Sort Code],	[Bank State],	[Flag Code],	'',	[Block_House No],	[Premises_Estate],	'17530101',	[ID Issue Date],	[Place of Issue],	[Bank Address],	[Bank Telephone No],	
		'17530101',	'0',	'',	[Employer City],	[E_Mail 2],	'0',	[Employer Status],	'',	'',	[Residential Address],	[State],	'',	[Correspondence Address],	[Correspondence Town],	[Correspondence Country],	[Upload Date],	[Date of Confirmation],	[Date of Last Promotion],	[Total Pensionable Salary],	[Salary Structure],	'',	0,	0,	[Salary Scale],	[Correspondence Town2],	0,	[Basic Pay],	[Employer Business Name],	[Industry Code],	[Employee Monthly Contribution],	[Employer Monthly Contribution],	[Total Monthly Contribution],	[Total Mon_ Pensionable Salary],	[Account Type],	[Employee Status],	[Employer State Code],	[Checked Date],	[Investproduct1],	[SchemeId],	[Invest percentage],	[Employer Address1],	[PIN Batch],	[Old PFA Code],	[LGA Name],	[State Name],	[State Of Posting Name],	[State Of Origin Name],	[Gender],	0,	'',	0,	[PFA Code],	[Block RSA Statement],	[Block Retiree Statement],	[Military],	'',	[AVC Agent Code],	[Document Library No],	[KYC],	[No_]
 , '',0,0x0 ,'','' from EASYREGSERVER.DBO.[Temporary Client]  with (nolock) where P_I_N ='" + pin + "'and No_ = '" & no & "'"
        Dim vendords As New DataSet
        vendords = returndataset("ARM", query)
        If errorresp <> "" Then Return errorresp

        'query for nok 
        querynok = "insert into [dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Pension BeneficiariesX] (
[Vendor No],[Line No],[Relationship],[Name],[Address],[Residential Address],[LGA],[LGA Name],[Internation Mobile],[Telephone No],[Email],[Surname],[First Name],[Middle Name],[Title],[State of Origin],[City],[Country],[Form Number],[Mobile Phone],[Office Phone],[Signature],[Left Thumbprint],[Right Thumbprint],[Gender old],[Home Phone],[Other Name],[State Code],[Correspondence Address],[PIN Invalid],[Picture],[Gender],[Personal E-mail],[State Name],[Country Name])
select '" & vendorno & "',[Line No],[Relationship],[Name],[Address],[Residential Address],'','','',[Telephone No],[Email],[Surname],[First Name],[Middle Name],[Title],[State of Origin],[City],[Country],[Form Number],[Mobile Phone],[Office Phone],[Signature],[Left Thumbprint],[Right Thumbprint],[Gender old],[Home Phone],[Other Name],[State Code],[Correspondence Address],[PIN Invalid],[Picture],[Gender],[Personal E-mail],[State Name],[Country Name] from EASYREGSERVER.DBO.[Temporary NOK] where TEMPORARY_ID ='" & no & "'"

        'query for beneficiary
        queryben = "insert into ARM.[dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Beneficiary Details] ([Vendor No],[Relationship],[Residential Address],[LGA],[LGA Name],[Internation Mobile], [Email],[Surname],[First Name],[Middle Name],[Title],[State of Origin],[City],[Country],[Mobile Phone],[State Code],[Correspondence Address],[Gender],[Personal E-mail],[State Name],[Country Name])
select '" & vendorno & "',[Relationship],[Residential Address],'','','',[Email],[Surname],[First Name],[Middle Name],[Title],[State of Origin],[City],[Country],[Mobile Phone],[State Code],[Correspondence Address],[Gender],[Personal E-mail],[State Name],[Country Name]
from EASYREGSERVER.DBO.[Temporary Beneficiary Details] where TEMPORARY_ID ='" & no & "'"


        'query for vendor additional 
        queryvenadd = "insert into ARM.[dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Vendor Additional] ([Vendor No],[NIN],[BVN],[Religion],[ID Number],[Place of Birth],[Highest Qualification],[Previously Employed],[Zip code],[Preffered Correspondence Adres],[Employer Rate of Contribution],[Employee Rate of Contribution],[Previous Business Name],[Previous Business Address],[Previous Business Town],[Last CFI Date],[Proof of Identity],[Proof of Address],[Proof of Identity Name],[Proof of Address Name],[Onboarding Source],[Last Form Ref No],[Moved to Sharepoint],[Select],[Selected by])
select '" & vendorno & "',[NIN],[BVN],[Religion],[ID Number],[Place of Birth],[Highest Qualification],[Previously Employed],[Zip code],[Preffered Correspondence Adres],[Employer Rate of Contribution],[Employee Rate of Contribution],[Previous Business Name],[Previous Business Address],[Previous Business Town],[Last CFI Date],[Proof of Identity],[Proof of Address],[Proof of Identity Name],[Proof of Address Name],[Onboarding Source],[Last Form Ref No],[Moved to Sharepoint],[Select],[Selected by]from EASYREGSERVER.DBO.[Temporary Client Additional] where TEMPORARY_ID ='" & no & "'"

        'query to update no_ series no.
        Dim noseriesquery As String = "update ARM.dbo.[" & My.MySettings.Default.COMPANYNAME.Trim & "$No_ Series Line] set [Last No_ Used] = CONVERT(VARCHAR(50),CAST([Last No_ Used] AS decimal) + 1 ) from [" & My.MySettings.Default.COMPANYNAME.Trim & "$No_ Series Line] where [Series Code] ='MEMBERSHIP'"
        Dim dst As New DataSet
        dst = returndataset("", querynok & " " & queryben & " " & queryvenadd & " " & noseriesquery)
        If errorresp <> "" Then Return errorresp
        Dim respmsg As String = MoveBiometrics("easyregserver", pin)
        If respmsg.ToString.ToLower = "success" Then
            query = "update EASYREGSERVER.DBO.[Temporary Client] set uploaded = 1 where No_ ='" & no & "' "
            returndataset("", query)
            Return vendorno
        End If
    End Function

    Dim errorresp As String = ""
    Private Function returndataset(ByVal databasenames As String, ByVal query As String) As DataSet
        errorresp = ""
        returndataset = Nothing
        Dim con1 As New SqlConnection
        Dim com1 As New SqlCommand
        Try
            con1 = New SqlConnection
            con1.ConnectionString = My.MySettings.Default.ARM_TESTINGConnectionString
            '   con1.ConnectionString = "Data Source=192.168.0.218;Initial Catalog=ARM;User ID=sa;Password=t@sting1"

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

            errorresp = ex.Message
            Return Nothing
        End Try
    End Function
    Private Function returnYYYYMMdd(ByVal dDate As String) As String
        '  Return (dDate)

        ' if to be used in a dll/webservice the above should be uncommented and the below should be commented

        dDate = dDate.Replace(" / ", "")
        dDate = dDate.Replace("/", "")
        dDate = dDate.Replace(" ", "")
        Dim mydate As String = Mid(dDate.Trim, 1, 2)
        Dim mymonth As String = Mid(dDate.Trim, 3, 2)
        Dim myyear As String = Mid(dDate.Trim, 5, 4)
        '  returnYYYYMMdd = (mymonth + " / " + mydate + " / " + myyear)
        returnYYYYMMdd = myyear + mymonth + mydate
    End Function
    <WebMethod()>
    Public Function RecordMassUpdateHelper(ByVal batchno As String, ByVal tbl As String, ByVal fld As String) As String
        Try
            RecordMassUpdateHelper = ""
            Dim mdt As New DataTable
            Dim IUSEDATE As Boolean = False
            Dim query As String
            Dim DS As New DataSet
            fld = Replace(fld, ".", "_")
            query = "Select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Mass Update Header] where [No_] = '" & batchno & "' select * from [" & My.MySettings.Default.COMPANYNAME.Trim & "$Mass Update Lines] where [No_] = '" & batchno & "' "
            DS = returndataset("", query)
            If DS Is Nothing Then Return "Connection Issue"
            If DS.Tables.Count = 0 Then Return "Communication Issue"
            If DS.Tables(0).Rows.Count = 0 Then Return "Invalid Batch No"
            If DS.Tables(1).Rows.Count = 0 Then Return "No lines for Header No"
            mdt = DS.Tables(1)
            Dim rds As DataSet
            If fld.ToUpper.Trim.Contains("DATE") Then IUSEDATE = True

            Dim tbno As Integer = mdt.Rows(0).Item("Table No_")
            If tbno = 23 Then
                For i As Integer = 0 To DS.Tables(1).Rows.Count - 1
                    If IUSEDATE = True Then mdt.Rows(i).Item("New Value") = returnYYYYMMdd(mdt.Rows(i).Item("New Value").ToString)
                    query = "Update [" & My.MySettings.Default.COMPANYNAME.Trim & "$" & tbl.Trim & "] set [" & fld & "]= '" & mdt.Rows(i).Item("New Value").ToString & "' where [No_] = '" & mdt.Rows(i).Item("Client No_").ToString & "'"
                    '  Return query
                    rds = returndataset("", query)

                    If errorresp = "" Then
                        query = "INSERT INTO [dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Change Log Entry] ([Date and Time],[Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value] " &
    " ,[Primary Key Field 2 No_],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID],[ERROR MSG],[SENT TO PENCOM],[LAST DATE SENT]) " &
        "  VALUES('" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DS.Tables(0).Rows(0).Item("Created By").ToString & "','" & mdt.Rows(i).Item("Table No_") & "','" & mdt.Rows(i).Item("Field No_") & "',1,'" & mdt.Rows(i).Item("Old Value").ToString & "','" & mdt.Rows(i).Item("New Value").ToString & "','Field1=0( " & mdt.Rows(i).Item("Client No_").ToString.Trim & ")',1,'" & mdt.Rows(i).Item("Client No_").ToString.Trim & "'" &
    " ,0,'',0,'',convert(VARBINARY(448), ''),'',0,'17530101') "
                        rds = returndataset("", query)
                    End If
                Next
            ElseIf tbno = 18
                For i As Integer = 0 To DS.Tables(1).Rows.Count - 1
                    query = "Update [" & My.MySettings.Default.COMPANYNAME.Trim & "$" & tbl.Trim & "] set [" & fld & "]= '" & mdt.Rows(i).Item("New Value").ToString & "' where [No_] = '" & mdt.Rows(i).Item("Employer No_").ToString & "'"
                    '  Return query

                    rds = returndataset("", query)

                    If errorresp = "" Then
                        query = "INSERT INTO [dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Change Log Entry] ([Date and Time],[Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value] " &
    " ,[Primary Key Field 2 No_],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID],[ERROR MSG],[SENT TO PENCOM],[LAST DATE SENT]) " &
        "  VALUES('" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DS.Tables(0).Rows(0).Item("Created By").ToString & "','" & mdt.Rows(i).Item("Table No_") & "','" & mdt.Rows(i).Item("Field No_") & "',1,'" & mdt.Rows(i).Item("Old Value").ToString & "','" & mdt.Rows(i).Item("New Value").ToString & "','Field1=0( " & mdt.Rows(i).Item("Employer No_").ToString.Trim & ")',1,'" & mdt.Rows(i).Item("Employer No_").ToString.Trim & "'" &
    " ,0,'',0,'',convert(VARBINARY(448), ''),'',0,'17530101') "
                        rds = returndataset("", query)
                    End If
                Next
            Else
                For i As Integer = 0 To DS.Tables(1).Rows.Count - 1
                    query = "Update [" & My.MySettings.Default.COMPANYNAME.Trim & "$" & tbl.Trim & "] set [" & fld & "]= '" & mdt.Rows(i).Item("New Value").ToString & "' where [Vendor No] = '" & mdt.Rows(i).Item("Client No_").ToString & "'"
                    '  Return query

                    rds = returndataset("", query)
                    If errorresp = "" Then

                        query = "INSERT INTO [dbo].[" & My.MySettings.Default.COMPANYNAME.Trim & "$Change Log Entry] ([Date and Time],[Time],[User ID],[Table No_],[Field No_],[Type of Change],[Old Value],[New Value],[Primary Key],[Primary Key Field 1 No_],[Primary Key Field 1 Value] " &
    " ,[Primary Key Field 2 No_],[Primary Key Field 2 Value],[Primary Key Field 3 No_],[Primary Key Field 3 Value],[Record ID],[ERROR MSG],[SENT TO PENCOM],[LAST DATE SENT]) " &
        "  VALUES('" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DateAdd(DateInterval.Hour, -1, Now) & "','" & DS.Tables(0).Rows(0).Item("Created By").ToString & "','" & mdt.Rows(i).Item("Table No_") & "','" & mdt.Rows(i).Item("Field No_") & "',1,'" & mdt.Rows(i).Item("Old Value").ToString & "','" & mdt.Rows(i).Item("New Value").ToString & "','Field1=0( " & mdt.Rows(i).Item("Client No_").ToString.Trim & ")',1,'" & mdt.Rows(i).Item("Client No_").ToString.Trim & "'" &
    " ,0,'',0,'',convert(VARBINARY(448), ''),'',0,'17530101') "
                        rds = returndataset("", query)
                    End If
                Next
            End If

            If errorresp <> "" Then Return errorresp
            Return RecordMassUpdateHelper
        Catch ex As Exception
            Return ex.Message
        End Try
    End Function

End Class