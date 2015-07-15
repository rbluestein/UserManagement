Imports System.Web
Imports System.Web.SessionState

Public Class Global
    Inherits System.Web.HttpApplication

#Region " Component Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application is started
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session is started

        Dim Enviro As New Enviro
        Session("Enviro") = Enviro

        Dim CompanyWorklistSession As New CompanyWorklistSession
        Session("CompanyWorklistSession") = CompanyWorklistSession

        Dim CarrierWorklistSession As New CarrierWorklistSession
        Session("CarrierWorklistSession") = CarrierWorklistSession

        Dim ClientWorklistSession As New ClientWorklistSession
        Session("ClientWorklistSession") = ClientWorklistSession

        Dim UserWorklistSession As New UserWorklistSession
        Session("UserWorklistSession") = UserWorklistSession

        Dim LicenseWorklistSession As New LicenseWorklistSession
        Session("LicenseWorklistSession") = LicenseWorklistSession

        Dim BulkLicSession As New BulkLicSession
        Session("BulkLicSession") = BulkLicSession

        Dim BulkApptSession As New BulkApptSession
        Session("BulkApptSesson") = BulkApptSession

        Dim BulkAppointmentDataPack As New BulkAppointmentDataPack
        Session("BulkAppointmentDataPack") = BulkAppointmentDataPack

        'Dim BrowserNav As New BrowserNav
        'Session("BrowserNav") = BrowserNav

        Dim QueryEnrollerSession As New QueryEnrollerSession
        Session("QueryEnrollerSession") = QueryEnrollerSession

        Dim QueryCarrierSession As New QueryCarrierSession
        Session("QueryCarrierSession") = QueryCarrierSession

        Dim QueryLicenseSession As New QueryLicenseSession
        Session("QueryLicenseSession") = QueryLicenseSession
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires at the beginning of each request
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires upon attempting to authenticate the use
    End Sub

    'Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
    '    ' Fires when an error occurs
    '    Dim HeaderMessage As String
    '    Dim msg As String
    '    Dim ex As Exception
    '    'Dim Request As HttpRequest

    '    ex = Server.GetLastError.InnerException
    '    'If ex.Message <> "Thread was being aborted." Then
    '    msg = ex.InnerException.Message & "~" & ex.InnerException.StackTrace
    '    msg = Replace(msg, vbCrLf, "~")
    '    Dim Report as New Report
    '    HeaderMessage = AppAdmin.Report(Report.ReportTypeEnum.Error, msg)
    '    ' Response.Redirect("ErrorPage.aspx?ErrorMsg=" & msg)
    '    Response.Redirect("ErrorPage.aspx?ErrorMsg=" & msg & "&HeaderMessage=" & HeaderMessage)
    '    'End If
    'End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        Dim Querystring As String
        Dim ErrorObj As ErrorObj
        Dim ErrorMessage As String
        Dim HeaderMessage As String
        Dim Report As New Report
        Dim ReportType As ReportTypeEnum

        ErrorObj = New ErrorObj(Server.GetLastError, Request)
        ErrorMessage = ErrorObj.GetDisplayErrorMessage()

        ReportType = ReportTypeEnum.Error
        If ErrorObj.IsCookieError Then
            ErrorMessage = "No cookie present."
        ElseIf ErrorObj.IsTimeOutError Then
            ReportType = ReportTypeEnum.Timeout
            ErrorMessage = "Application has timed out. Please close the application and log back in."
        ElseIf ErrorObj.IsPostTimeOutError Then
            ReportType = ReportTypeEnum.Timeout
            ErrorMessage = "Application has timed out. Please close the application and log back in."
        End If

        HeaderMessage = Report.Report(ReportType, ErrorMessage)

        Select Case ReportType
            Case ReportTypeEnum.Error, ReportTypeEnum.Timeout
                ErrorMessage = Replace(ErrorMessage, "#", "[sharp]")
                ErrorMessage = Replace(ErrorMessage, vbCrLf, "~")
                Response.Redirect("ErrorPage.aspx?ErrorMsg=" & ErrorMessage & "&HeaderMessage=" & HeaderMessage)
        End Select

    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the session ends
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Fires when the application ends
    End Sub

End Class
