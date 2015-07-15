Public Class calendar
    Inherits System.Web.UI.Page

    Protected PageCaption As HtmlGenericControl

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

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
        If Request.QueryString("targetctl") <> Nothing Then
            If Request.QueryString("targetctl") = "hdTermDate" Then
                PageCaption.InnerHtml = "Select Term Date"
            Else
                PageCaption.InnerHtml = "Select Date"
            End If
        Else
            PageCaption.InnerHtml = "Select Date"
        End If
    End Sub

End Class
