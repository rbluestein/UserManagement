Imports System.Data.SqlClient
Imports System.Web.Mail

Public Class EmailTest
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lblTo As System.Web.UI.WebControls.Label
    Protected WithEvents txtTo As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtBody As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblBody As System.Web.UI.WebControls.Label
    Protected WithEvents chkAttachment As System.Web.UI.WebControls.CheckBox
    Protected WithEvents lblAttachment As System.Web.UI.WebControls.Label
    Protected WithEvents btnSendMail As System.Web.UI.WebControls.Button
    Protected WithEvents txtResults As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblResults As System.Web.UI.WebControls.Label

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Coll As New Collection
        Dim SendMailResults As Results
        Dim AppAdmin As New AppAdmin


        If Page.IsPostBack Then
            txtResults.Text = String.Empty
            If txtTo.Text.Length > 0 Then
                If chkAttachment.Checked Then
                    Coll.Add("Who Is A Big Dog")
                End If
                SendMailResults = AppAdmin.SendEmail(txtTo.Text, "rbluestein@benefitvision.com;rhlavac@benefitvision.com", "", "Test: " & Date.Now, txtBody.Text, Coll)
                ' SendMailResults = SendEmail("rbluestein@benefitvision.com", Enviro.LoggedInUserID & "@benefitvision.com", "", "UserManagement error - " & System.Environment.UserName & "/" & System.Environment.MachineName, "An error occurred in the execution of UserManagement. Attached please find a copy of the application log and a screen shot.", Coll)
                If SendMailResults.Success Then
                    txtResults.Text = "Email sent"
                Else
                    txtResults.Text = "Email not sent" & vbCrLf & SendMailResults.Msg
                End If
            End If

        Else
            txtTo.Text = "rbluestein@benefitvision.com"
            txtBody.Text = "Have a nice day."
        End If
    End Sub
End Class
