Imports System

Public Class Response
    'INSTANT VB NOTE: The variable url was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private url_Renamed As String

    'INSTANT VB NOTE: The variable responseMessage was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private responseMessage_Renamed As String

    'INSTANT VB NOTE: The variable indicator was renamed since Visual Basic does not allow variables and other class members to have the same name:
    Private indicator_Renamed As Boolean

    Public Property Indicator() As Boolean
        Get
            Return Me.indicator_Renamed
        End Get
        Set(ByVal value As Boolean)
            Me.indicator_Renamed = value
        End Set
    End Property
    Public Property ResponseMessage() As String
        Get
            Return Me.responseMessage_Renamed
        End Get
        Set(ByVal value As String)
            Me.responseMessage_Renamed = value
        End Set
    End Property

    Public Property Url() As String
        Get
            Return Me.url_Renamed
        End Get
        Set(ByVal value As String)
            Me.url_Renamed = value
        End Set
    End Property

    Public Sub New()
    End Sub
End Class

