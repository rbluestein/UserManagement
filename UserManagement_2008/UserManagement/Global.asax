<%@ Application Language="VB" %>

<script runat="server">

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application startup
    End Sub
    
    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs on application shutdown
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


    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' Code that runs when a session ends. 
        ' Note: The Session_End event is raised only when the sessionstate mode
        ' is set to InProc in the Web.config file. If session mode is set to StateServer 
        ' or SQLServer, the event is not raised.
    End Sub
       
</script>