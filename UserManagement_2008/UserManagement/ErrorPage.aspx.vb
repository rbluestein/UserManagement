Partial Class ErrorPage
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Common As New Common
        Dim ErrorMessage As String = String.Empty

        ErrorMessage = Request.QueryString("ErrorMessage")
        ErrorMessage = Replace(ErrorMessage, "[sharp]", "#")
        ErrorMessage = Replace(ErrorMessage, "~", "<br>")
        litError.Text = Request.QueryString("HeaderMessage") & "<br><br>" & ErrorMessage
        PageCaption.InnerHtml = Common.GetPageCaption
    End Sub
End Class
