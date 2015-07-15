Partial Class LittleMessage
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblMessage.Text = Request.QueryString("Message")
    End Sub
End Class
