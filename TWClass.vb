Imports Newtonsoft.Json
Imports Newtonsoft.Json.Converters

Namespace QuickType
    Partial Public Class ProvBalheader
        <JsonProperty("userid")>
        Public Property userid As String

        <JsonProperty("password")>
        Public Property password As String

        <JsonProperty("token")>
        Public Property token As String
        <JsonProperty("quarter")>
        Public Property quarter As String

    End Class
    Partial Public Class ProvbalBody

        <JsonProperty("provisionalDetails")>
        Public Property provisionalDetails As provisionalDetail
    End Class


    Partial Public Class provisionalDetail

        <JsonProperty("referenceId")>
        Public Property referenceId As String
        <JsonProperty("rsaBalance")>
        Public Property rsaBalance As String
        <JsonProperty("asAtDate")>
        Public Property asAtDate As String
    End Class


    Public Class ProvObject
        Public Property provisionalDetails As List(Of provisionalDetail)
    End Class

    Partial Public Class thSummary

        <JsonProperty("employerCode")>
        Public Property employerCode As String
        <JsonProperty("firstname")>
        Public Property firstname As String
        <JsonProperty("fundCode")>
        Public Property fundCode As String
        <JsonProperty("middlename")>
        Public Property middlename As String
        <JsonProperty("quarterId")>
        Public Property quarterId As String
        <JsonProperty("referenceId")>
        Public Property referenceId As String
        <JsonProperty("rsapin")>
        Public Property rsapin As String
        <JsonProperty("rsaBalance")>
        Public Property rsaBalance As String
        <JsonProperty("surname")>
        Public Property surname As String

        <JsonProperty("tpfacode")>
        Public Property tpfacode As String
        <JsonProperty("ttlGainOrLoss")>
        Public Property ttlGainOrLoss As String
        <JsonProperty("ttlNoOfUnits")>
        Public Property ttlNoOfUnits As String
        <JsonProperty("unitPrice")>
        Public Property unitPrice As String
    End Class
    Partial Public Class detailRecord

        <JsonProperty("emplContribution")>
        Public Property emplContribution As String
        <JsonProperty("employeeContribution")>
        Public Property employeeContribution As String
        <JsonProperty("fees")>
        Public Property fees As String
        <JsonProperty("netContribution")>
        Public Property netContribution As String
        <JsonProperty("numberOfUnits")>
        Public Property numberOfUnits As String
        <JsonProperty("others")>
        Public Property others As String
        <JsonProperty("paymentDate")>
        Public Property paymentDate As String
        <JsonProperty("quarterId")>
        Public Property quarterId As String
        <JsonProperty("referenceId")>
        Public Property referenceId As String

        <JsonProperty("relatedMnthEnd")>
        Public Property relatedMnthEnd As String
        <JsonProperty("relatedMnthStart")>
        Public Property relatedMnthStart As String
        <JsonProperty("relatedPfaCode")>
        Public Property relatedPfaCode As String
        <JsonProperty("serialNo")>
        Public Property serialNo As String
        <JsonProperty("totalContribution")>
        Public Property totalContribution As String
        <JsonProperty("transactionType")>
        Public Property transactionType As String
        <JsonProperty("voluntaryContingent")>
        Public Property voluntaryContingent As String
        <JsonProperty("voluntaryRetirement")>
        Public Property voluntaryRetirement As String
        <JsonProperty("withdrawal")>
        Public Property withdrawal As String
    End Class
    Public Class thObject
        Public Property thSummary As thSummary
        Public Property detailRecords As List(Of detailRecord)
    End Class

End Namespace
