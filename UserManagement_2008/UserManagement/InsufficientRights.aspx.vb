
Partial Class InsufficientRights
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Common As New Common
        PageCaption.Text = Common.GetPageCaption
    End Sub
End Class
