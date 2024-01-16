Imports System.IO

Public Class ReconProcess
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        writefailure("")
        If IsPostBack Then Exit Sub
        If Request.QueryString("NO") <> Nothing Then
            txtno.Text = Request.QueryString("NO")
            Dim msg As String = ""
            lblmsg.Text = "Running Reconciliation process on NAV DB for " + txtno.Text + ". Please Wait..........."
            Dim ReconEngineServ As New ReconEngineServ
            msg = ReconEngineServ.ReconEngines(txtno.Text)
            If msg = "" Then
                writesuccess("Recon Header No : " + txtno.Text + " Processed Successfully" + vbCrLf + "Kindly close this page.")
                Response.Write("<script language='javascript'>self.close();</script>")
            Else
                writefailure(msg)
            End If
        End If
    End Sub

    Private Sub writefailure(ByVal msg As String)
        lblmsg.Text = msg
        lblmsg.ForeColor = Drawing.Color.Red
    End Sub
    Private Sub writesuccess(ByVal msg As String)
        lblmsg.Text = msg
        lblmsg.ForeColor = Drawing.Color.Green
    End Sub

End Class