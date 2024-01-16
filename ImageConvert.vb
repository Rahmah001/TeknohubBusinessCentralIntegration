Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Drawing
Imports PdfToImage
Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing.Imaging

Class Pdf2Jpg
    Public Function convertNIN(ByVal tempid As String, ByVal PdfFile As String) As String
        'PdfFile = "C:\\Users\\rahmo\\Desktop\\reference letter.pdf"
        Try
            Dim use2 As Boolean
            Dim PngFile As String = Replace(PdfFile.ToUpper, ".PDF", ".JPG")
            Dim PngFile2 As String = Replace(PngFile, "CR", "CL")
            Try
                Dim Conversion As List(Of String) = cs_pdf_to_image.Pdf2Image.Convert(PdfFile, PngFile)

            Catch ex As Exception
                use2 = True
                Try
                    File.Copy(PngFile, PngFile2)
                    Dim Conversion As List(Of String) = cs_pdf_to_image.Pdf2Image.Convert(PdfFile, PngFile2)

                Catch ex1 As Exception

                End Try

            End Try
            Dim Output As Bitmap
            If use2 = True Then Output = New Bitmap(PngFile2) Else Output = New Bitmap(PngFile)

            Dim QUERY As String = "UPDATE [TEMPORARY biometrics] SET NIN_CONSENT = @CONSENT WHERE TEMPORARY_ID = '" & tempid & "'"
            Dim cn As New SqlConnection
            Dim cmd As New SqlCommand
            cn.ConnectionString = My.MySettings.Default.EasyRegconnect


            cn.Open()


            Dim img As Byte() = ImageToByte(Output)

            cmd = New SqlCommand(QUERY, cn)
            cmd.Parameters.Add(New SqlParameter("@CONSENT", img))

            cmd.ExecuteNonQuery()
            cn.Close()
            Output.Dispose()
        Catch ex As Exception
            Return ex.Message

        End Try
        Return "Success"


    End Function
    Public Function getNINJpeg(ByVal no As String) As String
        Dim ds As New DataSet
        Dim penin As New PencomIntegrator
        Dim names As String = ""
        Dim address As String = ""
        Dim sign As Image
        Dim signbyte() As Byte
        Dim imgCon As New ImageConverter()

        'Try
        Dim sqlstring As String = "select  * from [Temporary Client] where No_ = '" & no & "' SELECT * FROM [Temporary Biometrics] where TEMPORARY_ID ='" & no & "'"
        ds = penin.returnEASYREGREC(sqlstring)
        If ds.Tables.Count < 2 Then Return ("Invalid record for temporary client / Biometrics")
        names = ds.Tables(0).Rows(0).Item("Surname").ToString + " " + ds.Tables(0).Rows(0).Item("first name").ToString + " " + ds.Tables(0).Rows(0).Item("middle names").ToString
        address = ds.Tables(0).Rows(0).Item("RESIDENTIAL_ADDRESS_NO").ToString + " " + ds.Tables(0).Rows(0).Item("Address").ToString
        signbyte = ds.Tables(1).Rows(0).Item("SIGNATURE")
        If signbyte Is Nothing Then Return ("Signature not available for the client")

        'Create a stream in memory containing the bytes that comprise the image.
        Using stream As New IO.MemoryStream(signbyte)
            'Read the stream and create an Image object from the data.'
            sign = Image.FromStream(stream)
        End Using

        Dim logoimg As Image
        Dim logobyte As Byte()
        logoimg = My.Resources.LOGO
        logobyte = ImageToByte(logoimg)


        Dim bm As New Bitmap(1000, 500)

        Dim gr As Graphics = Graphics.FromImage(bm)
        gr.Clear(Color.White)
        gr.DrawString("CUSTOMER AUTHORIZATION FOR ACCESS TO NATIONAL IDENTITY NUMBER (NIN) INFORMATION", SystemFonts.CaptionFont, Brushes.Black, New Point(1, 110))
        ' Save the result as a JPEG file.
        gr.DrawString("I hereby certify that the information provided in this form Is correct.", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 140))
        gr.DrawString(" I further consent and authorize the Natioanl Identity Management Comission (NIMC) to release my NIN information to the National Pension commission(PENCOM) upon request.", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 170))
        gr.DrawString("It is my understanding that Pencom shall excercise due care to ensure that my information ss secure and protected", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 200))
        gr.DrawString("Form Reference Number: " & no, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 220))
        gr.DrawString("Name: " & names, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 250))
        gr.DrawString("Address: " & address, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 280))
        gr.DrawString("Date: " & Today.ToShortDateString, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 310))
        If sign Is Nothing = False Then gr.DrawImage(sign, New PointF(260, 260))
        If logoimg Is Nothing = False Then gr.DrawImage(logoimg, New PointF(10, 0))

        Try
            ' bm.Save("c:\temp\test.jpg", ImageFormat.Jpeg)
            Dim ninarray As Byte()
            ninarray = ConvertToByteArray(bm)
            ' Dim objImage As Image = CType(bm, Image)
            'Using mStream As New MemoryStream()
            '    objImage.Save(mStream, objImage.RawFormat)
            '    ninarray = mStream.ToArray()
            'End Using



            Dim QUERY As String = "UPDATE [TEMPORARY biometrics] SET NIN_CONSENT = @nin WHERE TEMPORARY_ID = '" & no & "'"

            ' Query = "update [Temporary Client] set "
            ' QUERY += "Nin_Consent =@nin where No_= '" & no & "'"
            Dim cmdcom As New SqlCommand
            Dim smscon As New SqlConnection
            smscon.ConnectionString = My.MySettings.Default.EasyRegconnect
            smscon.Open()
            cmdcom = New SqlCommand(QUERY, smscon)
            'cmdcom.Parameters.Add("@nin", SqlDbType.Image).Value = DirectCast(imgCon.ConvertTo(objImage, GetType(Byte())), Byte())
            cmdcom.Parameters.Add("@nin", SqlDbType.Image).Value = ninarray
            cmdcom.ExecuteNonQuery()
            smscon.Close()
            smscon.Dispose()
            ' Return DirectCast(imgCon.ConvertTo(objImage, GetType(Byte())), Byte())

            Return "success"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Shared Function ConvertToByteArray(ByVal value As Bitmap) As Byte()
        Dim bitmapBytes As Byte()

        Using stream As New System.IO.MemoryStream

            value.Save(stream, ImageFormat.Jpeg)
            bitmapBytes = stream.ToArray

        End Using

        Return bitmapBytes

    End Function
    Public Function getNINJpegCFI(ByVal no As String) As String
        Dim ds As New DataSet
        Dim penin As New PencomIntegrator
        Dim names As String = ""
        Dim address As String = ""
        Dim Rrcode As String = ""
        Dim sign As Image
        Dim signbyte() As Byte
        Dim imgCon As New ImageConverter()

        'Try
        Dim sqlstring As String = "select  * from EasyRegServer.dbo.Enrollee where CODE = '" & no & "' SELECT * FROM EasyRegServer.dbo.Enrollee_Images where code ='" & no & "'"
        ds = penin.returnEASYREGREC(sqlstring)
        If ds.Tables.Count < 2 Then Return ("Invalid record for enrollee / Images")
        names = ds.Tables(0).Rows(0).Item("Last_Name").ToString + " " + ds.Tables(0).Rows(0).Item("First_Name").ToString + " " + ds.Tables(0).Rows(0).Item("Middle_Name").ToString
        address = ds.Tables(0).Rows(0).Item("RESIDENTIAL_ADDRESS_NO").ToString + " " + ds.Tables(0).Rows(0).Item("Residential_Address").ToString + ", " + ds.Tables(0).Rows(0).Item("Residential_AD_Town").ToString
        Rrcode = ds.Tables(0).Rows(0).Item("recapture_code").ToString
        signbyte = ds.Tables(1).Rows(0).Item("SIGNATURE")
        If Rrcode <> "" Then Rrcode = "33RR" + Rrcode.PadLeft(8, "0")
        If signbyte Is Nothing Then Return ("Signature not available for the client")

        'Create a stream in memory containing the bytes that comprise the image.
        Using stream As New IO.MemoryStream(signbyte)
            'Read the stream and create an Image object from the data.'
            sign = Image.FromStream(stream)
        End Using

        Dim logoimg As Image
        Dim logobyte As Byte()
        logoimg = My.Resources.LOGO
        logobyte = ImageToByte(logoimg)


        Dim bm As New Bitmap(1000, 500)

        Dim gr As Graphics = Graphics.FromImage(bm)
        gr.Clear(Color.White)
        gr.DrawString("CUSTOMER AUTHORIZATION FOR ACCESS TO NATIONAL IDENTITY NUMBER (NIN) INFORMATION", SystemFonts.CaptionFont, Brushes.Black, New Point(1, 110))
        ' Save the result as a JPEG file.
        gr.DrawString("I hereby certify that the information provided in this form Is correct.", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 140))
        gr.DrawString(" I further consent and authorize the Natioanl Identity Management Comission (NIMC) to release my NIN information to the National Pension commission(PENCOM) upon request.", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 170))
        gr.DrawString("It is my understanding that Pencom shall excercise due care to ensure that my information ss secure and protected", SystemFonts.DefaultFont, Brushes.Black, New Point(1, 200))
        gr.DrawString("Form Reference Number: " & Rrcode, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 220))
        gr.DrawString("Name: " & names, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 250))
        gr.DrawString("Address: " & address, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 280))
        gr.DrawString("Date: " & Today.ToShortDateString, SystemFonts.DefaultFont, Brushes.Black, New Point(1, 310))
        If sign Is Nothing = False Then gr.DrawImage(sign, New PointF(260, 260))
        If logoimg Is Nothing = False Then gr.DrawImage(logoimg, New PointF(10, 0))

        Try
            'bm.Save("c:\temp\test.jpg", ImageFormat.Jpeg)

            Dim objImage As Image = CType(bm, Image)



            Dim QUERY As String = "UPDATE [ENROLLEE_IMAGES] SET NIN_CONSENT = @nin WHERE CODE = '" & no & "'"

            ' Query = "update [Temporary Client] set "
            ' QUERY += "Nin_Consent =@nin where No_= '" & no & "'"
            Dim cmdcom As New SqlCommand
            Dim smscon As New SqlConnection
            smscon.ConnectionString = My.MySettings.Default.EasyRegconnect
            smscon.Open()
            cmdcom = New SqlCommand(QUERY, smscon)
            cmdcom.Parameters.Add("@nin", SqlDbType.Image).Value = DirectCast(imgCon.ConvertTo(objImage, GetType(Byte())), Byte())
            cmdcom.ExecuteNonQuery()
            smscon.Close()
            smscon.Dispose()
            ' Return DirectCast(imgCon.ConvertTo(objImage, GetType(Byte())), Byte())

            Return "success"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
    Public Function convertNINRecapture(ByVal tempid As String, ByVal PdfFile As String) As String
        'PdfFile = "C:\\Users\\rahmo\\Desktop\\reference letter.pdf"
        Dim mainerror As String = ""
        Try
            Dim use2 As Boolean
            Dim PngFile As String = Replace(PdfFile.ToUpper, ".PDF", ".JPG")
            Dim PngFile2 As String = Replace(PngFile, "AIC", "A1C")
            Try
                Dim Conversion As List(Of String) = cs_pdf_to_image.Pdf2Image.Convert(PdfFile, PngFile)

            Catch ex As Exception
                use2 = True
                Try
                    File.Copy(PngFile, PngFile2)
                    Dim Conversion As List(Of String) = cs_pdf_to_image.Pdf2Image.Convert(PdfFile, PngFile2)

                Catch ex1 As Exception
                    mainerror += ex1.Message + "|"
                End Try
                mainerror += ex.Message
            End Try
            Dim Output As Bitmap
            If use2 = True Then Output = New Bitmap(PngFile2) Else Output = New Bitmap(PngFile)

            Dim QUERY As String = "UPDATE [enrollee_Images] SET NIN_CONSENT = @CONSENT WHERE CODE = '" & tempid & "'"
            Dim cn As New SqlConnection
            Dim cmd As New SqlCommand
            cn.ConnectionString = My.MySettings.Default.EasyRegconnect


            cn.Open()


            Dim img As Byte() = ImageToByte(Output)

            cmd = New SqlCommand(QUERY, cn)
            cmd.Parameters.Add(New SqlParameter("@CONSENT", img))

            cmd.ExecuteNonQuery()
            cn.Close()
            Output.Dispose()
        Catch ex As Exception
            Return ex.Message & mainerror

        End Try
        Return "Success"


    End Function
    Public Shared Function ImageToByte(ByVal img As Image) As Byte()
        Dim converter As ImageConverter = New ImageConverter()
        Return CType(converter.ConvertTo(img, GetType(Byte())), Byte())
    End Function

    Public Function SaveExistingNin(ByVal tempid As String, ByVal jpgfile As String) As String
        'PdfFile = "C:\\Users\\rahmo\\Desktop\\reference letter.pdf"

        Dim PngFile As String = jpgfile

        Try

            Dim Output As Bitmap
            Output = New Bitmap(PngFile)

            Dim QUERY As String = "UPDATE [TEMPORARY biometrics] SET NIN_CONSENT = @CONSENT WHERE TEMPORARY_ID = '" & tempid & "'"
            Dim cn As New SqlConnection
            Dim cmd As New SqlCommand
            cn.ConnectionString = My.MySettings.Default.EasyRegconnect


            cn.Open()


            Dim img As Byte() = ImageToByte(Output)

            cmd = New SqlCommand(QUERY, cn)
            cmd.Parameters.Add(New SqlParameter("@CONSENT", img))

            cmd.ExecuteNonQuery()
            cn.Close()
            Output.Dispose()
        Catch ex As Exception
            Return ex.Message

        End Try
        Return "Success"


    End Function

    Public Function SaveExistingNinRecapure(ByVal tempid As String, ByVal jpgfile As String) As String
        'PdfFile = "C:\\Users\\rahmo\\Desktop\\reference letter.pdf"

        Dim PngFile As String = jpgfile

        Try

            Dim Output As Bitmap
            Output = New Bitmap(PngFile)

            Dim QUERY As String = "UPDATE [enrollee_Images] SET NIN_CONSENT = @CONSENT WHERE CODE = '" & tempid & "'"
            Dim cn As New SqlConnection
            Dim cmd As New SqlCommand
            cn.ConnectionString = My.MySettings.Default.EasyRegconnect


            cn.Open()


            Dim img As Byte() = ImageToByte(Output)

            cmd = New SqlCommand(QUERY, cn)
            cmd.Parameters.Add(New SqlParameter("@CONSENT", img))

            cmd.ExecuteNonQuery()
            cn.Close()
            Output.Dispose()
        Catch ex As Exception
            Return ex.Message

        End Try
        Return "Success"


    End Function

End Class
