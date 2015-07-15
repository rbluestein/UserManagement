Public Class TestPage
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label

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
        'Put user code to initialize the page here
        'Dim p As System.Security.Principal.WindowsPrincipal
        'p = System.Threading.Thread.CurrentPrincipal
        'Label1.Text = "#1: " & p.Identity.Name

        'Dim daisy As String
        'daisy = HttpContext.Current.User.Identity.Name.ToString
        'Label2.Text = "#2: " & daisy

        'Dim dog As String
        'dog = Request.ServerVariables("AUTH_USER")
        'Label3.Text = "#3 :" & dog
    End Sub

End Class
