Imports System.Data.SqlClient

Public Class BulkAppointmentMessage
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected PageCaption As HtmlGenericControl
    Protected WithEvents litDG As System.Web.UI.WebControls.Literal
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack
#End Region

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim DG As DG

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right check
            cRights = New RightsClass(Page)
            Dim RightsRqd(0) As String
            RightsRqd.SetValue(cRights.LicenseEdit, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)

            cBulkAppointmentDataPack = Session("BulkAppointmentDataPack")

            DG = New DG("RecordKey", cCommon, cRights, True, "EmbeddedTableDef", "RecordKey")
            DG.AddDataBoundColumn("RecordKey", "RecordKey", Nothing, Nothing, False, Nothing, Nothing, Nothing)
            DG.AddDataBoundColumn("State", "State", "State", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("CarrierID", "CarrierID", "CarrierID", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddDataBoundColumn("SaveType", "SaveType", "Type", Nothing, True, Nothing, Nothing, "align='left'")
            DG.AddBooleanColumn("Success", "Success", "Success", Nothing, True, 1, "Yes", "No", Nothing, "align='left'")
            DG.AddDataBoundColumn("Comments", "Comments", "Comments", Nothing, True, Nothing, Nothing, "align='left'")

            litDG.Text = DG.GetText(cCommon.GetExtendedTable(cBulkAppointmentDataPack.DT), Nothing)
            PageCaption.InnerHtml = cCommon.GetPageCaption

        Catch ex As Exception
            Throw New Exception("Error #720: BulkAppointmentMessage Page_Load. " & ex.Message)
        End Try
    End Sub
End Class
