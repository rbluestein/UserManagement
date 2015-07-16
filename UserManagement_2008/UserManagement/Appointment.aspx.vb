Imports System.Data
Imports System.Data.SqlClient

Partial Class Appointment
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cSess As LicenseWorklistSession
    'Protected WithEvents PlaceHolder1 As System.Web.UI.WebControls.PlaceHolder
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Results As Results
        Dim RequestAction As RequestActionEnum
        Dim ResponseAction As ResponseActionEnum

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(RightsClass.LicenseView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("LicenseWorklistSession")

            ' ___ Initialize the BulkAppointmentDataPack (used to display JIT results)
            cBulkAppointmentDataPack = Session("BulkAppointmentDataPack")
            cBulkAppointmentDataPack.Init()

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseActionEnum.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("LicenseWorklist.aspx?CalledBy=Child")
            Else
                DisplayPage(ResponseAction)

                If cBulkAppointmentDataPack.DT.Rows.Count = 0 Then
                    If Not Results.Msg = Nothing Then
                        litMsg.Text = "<script type=""text/javascript"">alert('" & Results.Msg & "')</script>"
                    End If
                Else
                    litMsg.Text = "<script type=""text/javascript"">window.showModalDialog('BulkAppointmentMessage.aspx', null, 'dialogHeight:600px;dialogWidth:570px;help:no;status:yes;scroll:yes;location:no');</script>"
                End If

                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj(ex, "Error #550: Appointment Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function ExecuteRequestAction(ByVal RequestAction As RequestActionEnum) As Results
        Dim ValidationResults As New Results
        Dim SaveResults As New Results
        Dim MyResults As New Results

        Try

            Select Case RequestAction
                Case RequestActionEnum.ReturnToParent
                    MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
                    MyResults.Success = True

                Case RequestActionEnum.CreateNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayBlank

                Case RequestActionEnum.SaveNew
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then

                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then

                            If cBulkAppointmentDataPack.DT.Rows.Count > 0 Then
                                MyResults.ResponseAction = ResponseActionEnum.DisplayExisting
                            Else
                                MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
                            End If
                            MyResults.Msg = SaveResults.Msg

                        Else
                            MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                            MyResults.Msg = SaveResults.Msg
                        End If

                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestActionEnum.LoadExisting
                    MyResults.ResponseAction = ResponseActionEnum.DisplayExisting

                Case RequestActionEnum.SaveExisting
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then

                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then

                            If cBulkAppointmentDataPack.DT.Rows.Count > 0 Then
                                MyResults.ResponseAction = ResponseActionEnum.DisplayExisting
                            Else
                                MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
                            End If
                            MyResults.Msg = SaveResults.Msg

                        Else
                            MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                            MyResults.Msg = SaveResults.Msg
                        End If

                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        'MyResults.ResponseAction = ResponseAction.DisplayExisting
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestActionEnum.NoSaveNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew

                Case RequestActionEnum.NoSaveExisting
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting

                Case RequestActionEnum.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #551: Appointment ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestActionEnum) As Results
        Dim MyResults As New Results
        Dim ErrColl As New Collection
        'Dim dt As DataTable
        'Dim ApptEffDate As Date
        'Dim ApptExpDate As Date
        'Dim ApptExpires As Boolean
        'Dim ApptIsPending As Boolean
        'Dim ApptIsActual As Boolean
        'Dim OKToProceed As Boolean = True
        'Dim PassLicenseTest As Boolean

        Try

            Dim LicRules As New LicRules(cCommon)
            cCommon.ValidateDropDown(ErrColl, ddCarrierID, 1, "carrier not provided")
            cCommon.ValidateDropDown(ErrColl, ddStatusCode, 1, "status code not provided")

            If IsDate(txtEffectiveDate.Text) AndAlso ddStatusCode.SelectedValue = "P" Then
                cCommon.ValidateErrorOnly(ErrColl, "an pending appointment may not have an effective date")
            ElseIf (Not IsDate(txtEffectiveDate.Text)) AndAlso ddStatusCode.SelectedValue = "X" Then
                cCommon.ValidateErrorOnly(ErrColl, "an effective appointment must have an effective date")
            End If
            If IsDate(txtApplicationDate.Text) AndAlso ddStatusCode.SelectedValue = "X" Then
                cCommon.ValidateErrorOnly(ErrColl, "an effective appointment may not have an application date")
            End If


            'If txtEffectiveDate.Text.Length > 0 And txtAppointmentNumber.Text.Length = 0 Then
            '    cCommon.ValidateErrorOnly(ErrColl, "an appointment must have a appointment number")
            'End If

            ' v1.04 removed
            'LicRules.LicAppt_DateSequenceCheck(ErrColl, txtEffectiveDate, txtExpirationDate)


            '' ___ Is the date sequence correct?
            'If txtEffectiveDate.Text.Length > 0 AndAlso txtExpirationDate.Text.Length > 0 Then
            '    If CType(txtEffectiveDate.Text, Date) > CType(txtExpirationDate.Text, Date) Then
            '        cCommon.ValidateErrorOnly(ErrColl, "the effective date of the appointment falls after the expiration date of the appointment")
            '    End If
            'End If

            ' ___ No more than one carrier per state per enroller.
            LicRules.Appt_TestOneCarrierPerState(RequestAction, ErrColl, cSess, ddCarrierID.SelectedItem.Value)
            'If RequestAction = RequestAction.SaveNew Then
            '    dt = cCommon.GetDT("SELECT Count (*) From UserAppointments WHERE UserID = '" & cSess.UserID & "' AND CarrierID = '" & ddCarrierID.SelectedItem.Value & "'  AND State = '" & cSess.State & "'")
            '    If dt.Rows(0)(0) > 0 Then
            '        cCommon.ValidateErrorOnly(ErrColl, "this enroller already has an appointment record with this carrier in this state.")
            '    End If
            '    dt = Nothing
            'ElseIf RequestAction = RequestAction.SaveExisting Then
            '    ' No action
            'End If

            LicRules.Appt_AllowThisEdit(RequestAction, ErrColl, cSess.UserID, cSess.State, txtEffectiveDate.Text, txtExpirationDate.Text, ddStatusCode.SelectedValue)

            '' ___ Is appointment pending or actual?
            'If txtEffectiveDate.Text = String.Empty Then
            '    ApptIsPending = True
            'Else
            '    ApptIsActual = True
            '    ApptEffDate = CDate(txtEffectiveDate.Text)
            'End If

            '' ___ A pending appointment passes the test if it can be applied against an application license of an effective license.
            'If ApptIsPending AndAlso OKToProceed Then
            '    dt = cCommon.GetDT("SELECT Count (*) FROM UserLicenses WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "'")
            '    If dt.Rows(0)(0) > 0 Then
            '        PassLicenseTest = True
            '    End If
            '    dt = Nothing
            'End If

            '' ___ An actual appointment passes the test if it can be applied against an effective license.
            'If ApptIsActual AndAlso OKToProceed Then

            '    ' ___ Get the appointment effective and expiration dates
            '    ApptEffDate = CType(txtEffectiveDate.Text, Date)
            '    If txtExpirationDate.Text.Length > 0 Then
            '        ApptExpires = True
            '        ApptExpDate = CType(txtExpirationDate.Text, Date)
            '    End If

            '    ' ___ Test the appointment against the licenses.
            '    PassLicenseTest = IsAppointmentCovered(ApptEffDate, ApptExpDate, ApptExpires)

            '    If Not PassLicenseTest Then
            '        cCommon.ValidateErrorOnly(ErrColl, "unable to find a state license against which this appointment may be applied")
            '    End If
            'End If

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #552: Appointment IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestActionEnum) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results
        Dim dt As DataTable
        Dim CurStatusCode As String = Nothing
        Dim JITAppt As New JITAppt(Me)
        Dim QueryPack As DBase.QueryPack = Nothing

        Try

            'If IsDate(txtEffectiveDate.Text) Then
            '    NewStatusCode = "X"
            'Else
            '    NewStatusCode = "P"
            'End If

            If RequestAction = RequestActionEnum.SaveNew Then
                Sql.Append("INSERT INTO UserAppointments (UserID, State, ApplicationDate, EffectiveDate, ExpirationDate, CarrierID, AppointmentNumber, StatusCode, StatusCodeLastChangeDate, AddDate, ChangeDate)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(cSess.UserID, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(cSess.State, False, StringTreatEnum.SideQts_SecApost) & ", ")

                'If NewStatusCode = "X" Then
                If ddStatusCode.SelectedValue = "X" Then
                    Sql.Append("null, ")
                Else
                    Sql.Append(cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
                End If

                Sql.Append(cCommon.DateOutHandler(txtEffectiveDate.Text, True, True) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddCarrierID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtAppointmentNumber.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                'Sql.Append("'" & NewStatusCode & "', ")
                Sql.Append(cCommon.StrOutHandler(ddStatusCode.SelectedValue, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")

                QueryPack = cCommon.ExecuteNonQueryWithQuerypack(Sql.ToString)

            ElseIf RequestAction = RequestActionEnum.SaveExisting Then

                ' Key = userid + state + carrierid + apptnum
                dt = cCommon.GetDT("SELECT * From UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "' AND CarrierID = '" & cSess.CarrierID & "' AND AppointmentNumber = '" & cSess.AppointmentNumber & "'")

                Sql.Append("INSERT INTO UserAppointments (UserID, State, ApplicationDate, EffectiveDate, ExpirationDate, CarrierID, AppointmentNumber, StatusCode, StatusCodeLastChangeDate, AddDate, ChangeDate)")
                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(cSess.UserID, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(cSess.State, False, StringTreatEnum.SideQts_SecApost) & ", ")

                'If NewStatusCode = "X" Then
                If ddStatusCode.SelectedValue = "X" Then
                    Sql.Append("null, ")
                Else
                    Sql.Append(cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
                End If

                Sql.Append(cCommon.DateOutHandler(txtEffectiveDate.Text, True, True) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddCarrierID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtAppointmentNumber.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")

                'If IsDate(txtEffectiveDate.Text) Then
                '    NewStatusCode = "X"
                'Else
                '    NewStatusCode = "P"
                'End If
                'Sql.Append(cCommon.StrOutHandler(NewStatusCode, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddStatusCode.SelectedValue, False, StringTreatEnum.SideQts_SecApost) & ", ")

                'If NewStatusCode = dt.Rows(0)("StatusCode") Then
                If ddStatusCode.SelectedValue = dt.Rows(0)("StatusCode") Then
                    Sql.Append(cCommon.DateOutHandler(dt.Rows(0)("StatusCodeLastChangeDate"), False, True) & ", ")
                Else
                    Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                End If
                Sql.Append(cCommon.DateOutHandler(dt.Rows(0)("AddDate"), False, True) & ", ")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")


                QueryPack = cCommon.GetDTWithQueryPack("DELETE UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "' AND CarrierID = '" & cSess.CarrierID & "' AND AppointmentNumber = '" & cSess.AppointmentNumber & "'")
                If QueryPack.Success Then
                    QueryPack = cCommon.ExecuteNonQueryWithQuerypack(Sql.ToString)
                End If


                'dt = cCommon.GetDT("SELECT StatusCode From UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "' AND CarrierID = '" & cSess.CarrierID & "' AND AppointmentNumber = '" & cSess.AppointmentNumber & "'")
                'CurStatusCode = dt.Rows(0)("StatusCode")

                'Sql.Append("UPDATE UserAppointments Set ")
                'Sql.Append("State = " & cCommon.StrOutHandler(cSess.State, False, True) & ", ")
                'Sql.Append("ApplicationDate = " & cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
                'Sql.Append("EffectiveDate = " & cCommon.DateOutHandler(txtEffectiveDate.Text, True, True) & ", ")
                'Sql.Append("ExpirationDate = " & cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")
                'Sql.Append("CarrierID = " & cCommon.StrOutHandler(ddCarrierID.SelectedItem.Value, False, True) & ", ")
                'Sql.Append("AppointmentNumber = " & cCommon.StrOutHandler(txtAppointmentNumber.Text, False, True) & ", ")

                'If IsDate(txtEffectiveDate.Text) Then
                '    NewStatusCode = "X"
                'Else
                '    NewStatusCode = "P"
                'End If
                'Sql.Append("StatusCode = '" & NewStatusCode & "', ")

                'If CurStatusCode <> NewStatusCode Then
                '    Sql.Append("StatusCodeLastChangeDate = '" & cCommon.GetServerDateTime & "', ")
                'End If

                'Sql.Append("ChangeDate = '" & cCommon.GetServerDateTime & "' ")
                'Sql.Append(" WHERE UserID = '" & cSess.UserID & "' AND State ='" & cSess.State & "' AND CarrierID = '" & cSess.CarrierID & "'")
            End If

            If QueryPack.Success Then
                MyResults.Success = True
                MyResults.Msg = "Update complete."
                cSess.CarrierID = ddCarrierID.SelectedItem.Value
                cSess.EffectiveDate = txtEffectiveDate.Text
                cSess.AppointmentNumber = txtAppointmentNumber.Text
            Else
                MyResults.Success = False
                MyResults.Msg = QueryPack.TechErrMsg
            End If


            If QueryPack.Success Then
                If RequestAction = (RequestActionEnum.SaveNew AndAlso ddStatusCode.SelectedValue = "X") Or (RequestActionEnum.SaveExisting AndAlso CurStatusCode <> "X" AndAlso ddStatusCode.SelectedValue = "X") Then
                    If JITAppt.IsJITTrigger(cSess.UserID, cSess.State, cSess.CarrierID) Then
                        JITAppt.PerformJustInTimeAppointments(cSess.UserID, cSess.State, cSess.CarrierID, txtApplicationDate.Text, txtAppointmentNumber.Text, txtEffectiveDate.Text, txtExpirationDate.Text)
                        If cBulkAppointmentDataPack.DT.Rows.Count > 0 Then
                            cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.JITTrigger, cSess.State, cSess.CarrierID, 1, Nothing, True)
                        End If
                    End If
                End If
            End If

            Return MyResults

        Catch ex As Exception
            MyResults.Success = False
            MyResults.Msg = "Unable to save record."
            Return MyResults
        End Try

    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseActionEnum)
        Dim i As Integer
        Dim dt As DataTable
        Dim Sql As String
        Dim CarrierID As String
        Dim StatusCode As String

        Try

            ' ___ Heading/UserID
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.DisplayUserInputNew
                    litHeading.Text = "New Appointment"
                Case ResponseActionEnum.DisplayExisting, ResponseActionEnum.DisplayUserInputExisting
                    If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                        litHeading.Text = "Edit Appointment"
                    Else
                        litHeading.Text = "View Appointment"
                    End If
            End Select

            ' ___ Format the controls
            FormatControls(ResponseAction)

            ' ___ Enroller name heading
            Sql = "SELECT LastName + ', ' +  FirstName + ' ' + MI FullName FROM Users WHERE UserID ='" & cSess.UserID & "'"
            dt = cCommon.GetDT(Sql)
            lblEnrollerName.Text = "Name:&nbsp;&nbsp;" & dt.Rows(0)("FullName")

            If Not Page.IsPostBack Then

                ' ___ Carrier dropdown
                dt = cCommon.GetDT("SELECT CarrierID FROM Codes_CarrierID ORDER BY CarrierID")
                ddCarrierID.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    ddCarrierID.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                Next

                ' ___ StatusCode
                ddStatusCode.Items.Add(New ListItem("", ""))
                ddStatusCode.Items.Add(New ListItem("Pending", "P"))
                ddStatusCode.Items.Add(New ListItem("Effective", "X"))
                ddStatusCode.Items.Add(New ListItem("Term", "T"))

            End If

            txtState.Text = cSess.State

            ' ___ Get the data for ResponseAction.DisplayExisting only
            If ResponseAction = ResponseActionEnum.DisplayExisting Then
                'If cEffectiveDate.Length = 0 Then
                '    dt = Common.GetDT("SELECT * From UserAppointments Where UserID = '" & cSubjUserID & "' AND State = '" & cState & "' AND EffectiveDate is null AND CarrierID = '" & cCarrierID & "'")
                'Else
                '    dt = Common.GetDT("SELECT * From UserAppointments Where UserID = '" & cSubjUserID & "' AND State = '" & cState & "' AND EffectiveDate = '" & cEffectiveDate & "' AND CarrierID = '" & cCarrierID & "'")
                'End If

                '  dt = Common.GetDT("SELECT * From UserAppointments WHERE UserID = '" & Sess.UserID & "' AND State = '" & Sess.State & "' AND CarrierID = '" & Sess.CarrierID & "'")
                dt = cCommon.GetDT("SELECT * From UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "' AND CarrierID = '" & cSess.CarrierID & "' AND AppointmentNumber = '" & cSess.AppointmentNumber & "'")


                ddCarrierID.ClearSelection()
                CarrierID = dt.Rows(0)("CarrierID").ToLower
                For i = 0 To ddCarrierID.Items.Count - 1
                    If ddCarrierID.Items(i).Value.ToLower = CarrierID Then
                        ddCarrierID.SelectedIndex = i
                        Exit For
                    End If
                Next

                'ddCarrierID.Items.FindByValue(dt.Rows(0)("CarrierID")).Selected = True
                txtCarrierID.Text = cCommon.StrInHandler(dt.Rows(0)("CarrierID"))


                ddStatusCode.ClearSelection()
                For i = 0 To ddStatusCode.Items.Count - 1
                    If ddStatusCode.Items(i).Value = dt.Rows(0)("StatusCode") Then
                        ddStatusCode.SelectedIndex = i
                    End If
                Next
                StatusCode = dt.Rows(0)("StatusCode")
                Select Case StatusCode
                    Case String.Empty
                        txtStatusCode.Text = String.Empty
                    Case "P"
                        txtStatusCode.Text = "Pending"
                    Case "X"
                        txtStatusCode.Text = "Effective"
                    Case "T"
                        txtStatusCode.Text = "Term"
                End Select


                txtState.Text = cCommon.StrInHandler(dt.Rows(0)("State"))
                txtAppointmentNumber.Text = cCommon.StrInHandler(dt.Rows(0)("AppointmentNumber"))
                txtApplicationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ApplicationDate"))
                txtEffectiveDate.Text = cCommon.DateInHandler(dt.Rows(0)("EffectiveDate"))
                txtExpirationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ExpirationDate"))
                'chkJIT.Checked = dt.Rows(0)("JIT")
            End If

        Catch ex As Exception
            Throw New Exception("Error #554: Appointment DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls(ByVal ResponseAction As ResponseActionEnum)
        Try

            Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 300)

            ' ___ View/Edit
            If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"

                ddCarrierID.Visible = True
                txtCarrierID.Visible = False

                ddStatusCode.Visible = True
                txtStatusCode.Visible = False

                ' chkJIT.Visible = True
                'txtJIT.Visible = False
                lblApplicationDateLink.Visible = True
                lblEffectiveDateLink.Visible = True
                lblExpirationDateLink.Visible = True

                Style.AddStyle(txtAppointmentNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableWhite, 300)

            Else

                ddCarrierID.Visible = False
                Style.AddStyle(txtCarrierID, Style.StyleType.NoneditableGrayed, 300)


                ddStatusCode.Visible = False
                Style.AddStyle(txtStatusCode, Style.StyleType.NoneditableGrayed, 300)

                'chkJIT.Visible = False
                ' Style.AddStyle(txtJIT, Style.StyleType.NoneditableGrayed, 300)
                lblApplicationDateLink.Visible = False
                lblEffectiveDateLink.Visible = False
                lblExpirationDateLink.Visible = False

                Style.AddStyle(txtAppointmentNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableGrayed, 300)
            End If

        Catch ex As Exception
            Throw New Exception("Error #555: Appointment FormatControls. " & ex.Message)
        End Try
    End Sub
End Class
