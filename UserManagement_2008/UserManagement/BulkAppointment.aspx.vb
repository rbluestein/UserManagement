Imports System.Data
Imports System.Data.SqlClient

Public Class BulkAppointment
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    'Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    'Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    'Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    'Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Private cBulkApptSession As BulkApptSession
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack
    Private cJITAppt As JITAppt
#End Region


    Private Enum LocalRequestAction
        CreateNew = 1
        LoadExisting = 2
        SaveNew = 3
        SaveExisting = 4
        SaveNewOrExisting = 5
        NoSaveNew = 6
        NoSaveExisting = 7
        ReturnToParent = 8
        [Date] = 9
        Other = 10

        EnrollerLetterClick = 11
        EnrollerAdd = 12
        EnrollerRemove = 13
        CarrierAdd = 14
        CarrierRemove = 15
        StateAdd = 16
        StateRemove = 17
        StateSourceIndexChange = 18
        GoToLicense = 19
        'StatesPerCarrier = 20
        'CarriersPerState = 21
    End Enum

    Private Enum LocalResponseAction
        DisplayBlank = 1
        DisplayUserInputNew = 2
        DisplayUserInputExisting = 3
        DisplayUserInputNewOrExisting = 4
        DisplayExisting = 5
        ReturnToCallingPage = 6

        EnrollerLetterClick = 7
        EnrollerAdd = 8
        EnrollerRemove = 9
        CarrierAdd = 10
        CarrierRemove = 11
        StateAdd = 12
        StateRemove = 13
        StateSourceIndexChange = 14
        GoToLicense = 15
        'StatesPerCarrier = 16
        'CarriersPerState = 17
    End Enum

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Results As LocalResults
        Dim RequestAction As LocalRequestAction
        Dim ResponseAction As LocalResponseAction
        Dim PageMode As PageMode

        Try

            ' ___ Instantiate objects
            cEnviro = Session("Enviro")
            cCommon = New Common
            cJITAppt = New JITAppt(Me)

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(RightsClass.LicenseEdit, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cBulkApptSession = Session("BulkApptSesson")

            ' ___ Initialize the BulkAppointmentDataPack 
            cBulkAppointmentDataPack = Page.Session("BulkAppointmentDataPack")
            cBulkAppointmentDataPack.Init()

            ' ___ Get the page mode
            PageMode = cCommon.GetPageMode(Page, cBulkApptSession)

            ' ___ Get RequestAction
            RequestAction = GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(PageMode, RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = LocalResponseAction.ReturnToCallingPage Then
                Response.Redirect("LicenseWorklist.aspx?CalledBy=Child")
            Else
                DisplayPage(PageMode, ResponseAction)
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
            ' PageCaption.InnerHtml = cCommon.GetPageCaption
            PageCaption.Text = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj(ex, "Error #702: BulkAppointment Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function GetRequestAction(ByVal Page As Page) As LocalRequestAction
        Dim ActionType As String = Nothing
        Dim RequestAction As LocalRequestAction
        Dim hdResponseAction As String
        Dim hdAction As String
        Dim CallType As String

        Try

            If Page.Request.QueryString("CallType") = Nothing OrElse Page.Request.QueryString("CallType") = String.Empty Then
                CallType = String.Empty
            Else
                CallType = Page.Request.QueryString("CallType")
                CallType = CallType.ToLower
            End If

            If Page.Request.Form("hdAction") = Nothing OrElse Page.Request.Form("hdAction") = String.Empty Then
                hdAction = String.Empty
            Else
                hdAction = Page.Request.Form("hdAction")
                hdAction = hdAction.ToLower
            End If

            If Not Page.IsPostBack Then
                ActionType = "record"
            Else
                If hdAction = "enrollerletterclick" Then
                    ActionType = "selection"
                ElseIf hdAction = "enrolleradd" Then
                    ActionType = "selection"
                ElseIf hdAction = "enrollerremove" Then
                    ActionType = "selection"
                ElseIf hdAction = "carrieradd" Then
                    ActionType = "selection"
                ElseIf hdAction = "carrierremove" Then
                    ActionType = "selection"
                ElseIf hdAction = "stateadd" Then
                    ActionType = "selection"
                ElseIf hdAction = "stateremove" Then
                    ActionType = "selection"
                ElseIf hdAction = "gotolicense" Then
                    ActionType = "selection"
                ElseIf hdAction = "statesourceindexchange" Then
                    ActionType = "selection"
                ElseIf hdAction = "statespercarrier" Then
                    ActionType = "selection"
                ElseIf hdAction = "carriersperstate" Then
                    ActionType = "selection"
                ElseIf hdAction = "update" Then
                    ActionType = "record"
                ElseIf hdAction = "return" Then
                    ActionType = "return"
                ElseIf hdAction = "confirmation" Then
                    ActionType = "confirmation"
                ElseIf hdAction = "clientselectionchanged" Then
                    ActionType = "clientselectionchanged"
                End If
            End If

            Select Case ActionType
                Case "return"
                    RequestAction = LocalRequestAction.ReturnToParent

                Case "selection"
                    Select Case hdAction
                        Case "enrollerletterclick"
                            RequestAction = LocalRequestAction.EnrollerLetterClick
                        Case "enrolleradd"
                            RequestAction = LocalRequestAction.EnrollerAdd
                        Case "enrollerremove"
                            RequestAction = LocalRequestAction.EnrollerRemove
                        Case "carrieradd"
                            RequestAction = LocalRequestAction.CarrierAdd
                        Case "carrierremove"
                            RequestAction = LocalRequestAction.CarrierRemove
                        Case "stateadd"
                            RequestAction = LocalRequestAction.StateAdd
                        Case "stateremove"
                            RequestAction = LocalRequestAction.StateRemove
                        Case "statesourceindexchange"
                            RequestAction = LocalRequestAction.StateSourceIndexChange
                        Case "gotolicense"
                            RequestAction = LocalRequestAction.GoToLicense
                    End Select

                Case "record"
                    If Not Page.IsPostBack Then
                        Select Case CallType
                            Case "new", ""
                                RequestAction = LocalRequestAction.CreateNew
                            Case "existing"
                                RequestAction = LocalRequestAction.LoadExisting
                        End Select
                    Else
                        Select Case Page.Request.Form("hdResponseAction")
                            Case LocalResponseAction.DisplayBlank.ToString
                                RequestAction = LocalRequestAction.SaveNew
                            Case LocalResponseAction.DisplayExisting.ToString
                                RequestAction = LocalRequestAction.SaveExisting
                            Case LocalResponseAction.DisplayUserInputNew.ToString
                                RequestAction = LocalRequestAction.SaveNew
                            Case LocalResponseAction.DisplayUserInputExisting.ToString
                                RequestAction = LocalRequestAction.SaveExisting
                            Case LocalResponseAction.DisplayUserInputNewOrExisting.ToString
                                RequestAction = LocalRequestAction.SaveNewOrExisting
                        End Select
                    End If

                Case "confirmation"
                    hdResponseAction = Page.Request.Form("hdResponseAction")
                    If hdResponseAction = "DisplayBlank" Or hdResponseAction = "DisplayUserInputNew" Then
                        If Page.Request.Form("hdConfirm") = "yes" Then
                            RequestAction = LocalRequestAction.SaveNew
                        Else
                            RequestAction = LocalRequestAction.NoSaveNew
                        End If
                    ElseIf hdResponseAction = "DisplayExisting" Or hdResponseAction = "DisplayUserInputExisting" Then
                        If Page.Request.Form("hdConfirm") = "yes" Then
                            RequestAction = LocalRequestAction.SaveExisting
                        Else
                            RequestAction = LocalRequestAction.NoSaveExisting
                        End If
                    End If

                Case "clientselectionchanged"
                    RequestAction = LocalRequestAction.Other

            End Select

            Return RequestAction

        Catch ex As Exception
            Throw New Exception("Error #703: BulkAppointment GetRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function ExecuteRequestAction(ByVal PageMode As PageMode, ByVal RequestAction As LocalRequestAction) As LocalResults
        'Dim ValidationResults As New Results
        Dim SaveResults As New LocalResults
        Dim MyResults As New LocalResults

        Try

            Select Case RequestAction
                Case LocalRequestAction.ReturnToParent
                    MyResults.ResponseAction = LocalResponseAction.ReturnToCallingPage
                    MyResults.Success = True

                Case LocalRequestAction.CreateNew
                    MyResults.ResponseAction = LocalResponseAction.DisplayBlank

                Case LocalRequestAction.SaveNew
                    'ValidationResults = IsDataValid(RequestAction)
                    'If ValidationResults.Success Then
                    SaveResults = PerformSave(RequestAction)
                    If SaveResults.Success Then
                        MyResults.ResponseAction = LocalResponseAction.DisplayBlank
                        MyResults.Msg = SaveResults.Msg
                        'MyResults.Msg = GetBulkResults(SaveResults)
                    Else
                        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                        MyResults.Msg = SaveResults.Msg
                    End If
                    'ElseIf ValidationResults.ObtainConfirm Then
                    'MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    'MyResults.Msg = ValidationResults.Msg
                    'Else
                    'MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    'MyResults.Msg = ValidationResults.Msg
                    'End If

                Case LocalRequestAction.LoadExisting
                    MyResults.ResponseAction = LocalResponseAction.DisplayExisting

                Case LocalRequestAction.SaveExisting
                    'ValidationResults = IsDataValid(RequestAction)
                    'If ValidationResults.Success Then
                    SaveResults = PerformSave(RequestAction)
                    If SaveResults.Success Then
                        MyResults.ResponseAction = LocalResponseAction.DisplayBlank
                        MyResults.Msg = SaveResults.Msg
                        'MyResults.Msg = GetBulkResults(SaveResults)
                    Else
                        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                        MyResults.Msg = SaveResults.Msg
                    End If
                    'ElseIf ValidationResults.ObtainConfirm Then
                    'MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                    'MyResults.Msg = ValidationResults.Msg
                    'Else
                    'MyResults.ResponseAction = LocalResponseAction.DisplayExisting
                    'MyResults.Msg = ValidationResults.Msg
                    'End If

                Case LocalRequestAction.SaveNewOrExisting
                    'ValidationResults = IsDataValid(RequestAction)
                    'If ValidationResults.Success Then
                    SaveResults = PerformSave(RequestAction)
                    If SaveResults.Success Then
                        MyResults.ResponseAction = LocalResponseAction.DisplayBlank
                        MyResults.Msg = SaveResults.Msg
                        'MyResults.Msg = GetBulkResults(SaveResults)
                    Else
                        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                        MyResults.Msg = SaveResults.Msg
                    End If
                    'ElseIf ValidationResults.ObtainConfirm Then
                    'MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting
                    'MyResults.Msg = ValidationResults.Msg
                    'Else
                    'MyResults.ResponseAction = LocalResponseAction.DisplayExisting
                    'MyResults.Msg = ValidationResults.Msg
                    'End If

                Case LocalRequestAction.NoSaveNew
                    MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew

                Case LocalRequestAction.NoSaveExisting
                    MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting

                Case LocalRequestAction.EnrollerLetterClick
                    MyResults.ResponseAction = LocalResponseAction.EnrollerLetterClick

                Case LocalRequestAction.EnrollerAdd
                    MyResults.ResponseAction = LocalResponseAction.EnrollerAdd

                Case LocalRequestAction.EnrollerRemove
                    MyResults.ResponseAction = LocalResponseAction.EnrollerRemove

                Case LocalRequestAction.CarrierAdd
                    MyResults.ResponseAction = LocalResponseAction.CarrierAdd

                Case LocalRequestAction.CarrierRemove
                    MyResults.ResponseAction = LocalResponseAction.CarrierRemove

                Case LocalRequestAction.StateAdd
                    MyResults.ResponseAction = LocalResponseAction.StateAdd

                Case LocalRequestAction.StateRemove
                    MyResults.ResponseAction = LocalResponseAction.StateRemove

                Case LocalRequestAction.GoToLicense
                    HandleData(PageMode, False)
                    Response.Redirect("BulkLicense.aspx?CalledBy=Other")

                Case LocalRequestAction.StateSourceIndexChange
                    MyResults.ResponseAction = LocalResponseAction.StateSourceIndexChange

                Case LocalRequestAction.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #704: BulkAppointment ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    'Private Function GetBulkResults(ByRef SaveResults As BulkResults) As String
    '    Dim i As Integer
    '    Dim Msg As String = String.Empty
    '    With SaveResults.ReportColl
    '        For i = 1 To .Count
    '            Msg &= .Item(i) & "\n"
    '        Next
    '    End With
    '    Return Msg
    'End Function

    Private Function IsDataValid(ByVal RequestAction As LocalRequestAction) As LocalResults
        Dim MyResults As New LocalResults
        Dim ErrColl As New Collection

        Try

            Dim LicRules As New LicRules(cCommon)

            If txtEnrollerTarget.Text = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no enroller specified")
            End If
            If cBulkApptSession.CarrierTarget = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no carrier specified")
            End If
            If cBulkApptSession.StateTarget = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no state specified")
            End If

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #705: BulkAppointment IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As LocalRequestAction) As LocalResults
        Dim i, j As Integer
        Dim SaveResults As New LocalResults
        Dim ErrColl As New Collection

        Try

            Dim LicRules As New LicRules(cCommon)

            If txtEnrollerTarget.Text = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no enroller specified")
            End If
            If cBulkApptSession.CarrierTarget = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no carrier specified")
            End If
            If cBulkApptSession.StateTarget = String.Empty Then
                cCommon.ValidateErrorOnly(ErrColl, "no state specified")
            End If

            cCommon.ValidateDropDown(ErrColl, ddStatusCode, 1, "status code not provided")

            If IsDate(txtEffectiveDate.Text) AndAlso ddStatusCode.SelectedValue = "P" Then
                cCommon.ValidateErrorOnly(ErrColl, "an pending appointment may not have an effective date")
            ElseIf (Not IsDate(txtEffectiveDate.Text)) AndAlso ddStatusCode.SelectedValue = "X" Then
                cCommon.ValidateErrorOnly(ErrColl, "an effective appointment must have an effective date")
            End If
            If IsDate(txtApplicationDate.Text) AndAlso ddStatusCode.SelectedValue = "X" Then
                cCommon.ValidateErrorOnly(ErrColl, "an effective appointment may not have an application date")
            End If


            ' v1.04 removed
            'LicRules.LicAppt_DateSequenceCheck(ErrColl, txtEffectiveDate, txtExpirationDate)

            If ErrColl.Count > 0 Then
                SaveResults.Success = False
                SaveResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
                Return SaveResults
            End If

            'Select Case cBulkApptSession.PerStatePerCarrier
            '    Case BulkApptSession.PerStatePerCarrierEnum.CarriersPerState
            '        For i = 0 To cBulkApptSession.CarrierTargetList.Count - 1
            '            PerformSave2(LicRules, SaveResults, cBulkApptSession.StateTarget, cBulkApptSession.CarrierTargetList.GetKey(i))
            '        Next

            '    Case BulkApptSession.PerStatePerCarrierEnum.StatesPerCarrier
            '        For i = 0 To cBulkApptSession.StateTargetList.Count - 1
            '            PerformSave2(LicRules, SaveResults, cBulkApptSession.StateTargetList.GetKey(i), cBulkApptSession.CarrierTarget)
            '        Next
            'End Select

            For i = 0 To cBulkApptSession.StateTargetList.Count - 1
                For j = 0 To cBulkApptSession.CarrierTargetList.Count - 1
                    PerformSave2(LicRules, SaveResults, cBulkApptSession.StateTargetList.GetKey(i), cBulkApptSession.CarrierTargetList.GetKey(j))
                Next
            Next

            ' ___ When an appointment effective date is present, process as possible JIT appointment
            If IsDate(txtEffectiveDate.Text) Then
                For i = 0 To cBulkAppointmentDataPack.DT.Rows.Count - 1
                    If cBulkAppointmentDataPack.DT.Rows(i)("SaveType") = BulkAppointmentDataPack.SaveTypeClass.ApptExpressJITTrigger Then
                        cJITAppt.PerformJustInTimeAppointments(cBulkApptSession.UserId, cBulkAppointmentDataPack.DT.Rows(i)("State"), cBulkAppointmentDataPack.DT.Rows(i)("CarrierID"), txtApplicationDate.Text, txtAppointmentNumber.Text, txtEffectiveDate.Text, txtExpirationDate.Text)
                    End If
                Next
            End If

            Return SaveResults

        Catch ex As Exception
            Throw New Exception("Error #706: BulkAppointment PerformSave. " & ex.Message)
        End Try
    End Function

    Private Sub PerformSave2(ByRef LicRules As LicRules, ByRef SaveResults As LocalResults, ByVal State As String, ByVal CarrierID As String)
        Dim ApptIsCovered As Boolean
        Dim OrphanTestArgs As LicRules.OrphanTestArgs
        Dim dtLic As DataTable

        Try

            OrphanTestArgs = New LicRules.OrphanTestArgs
            dtLic = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate from UserLicenses WHERE  UserID='" & cBulkApptSession.UserId & "' AND  State = '" & State & "'")
            OrphanTestArgs.GetLicenseStatus(dtLic.Rows(0)("EffectiveDate"), dtLic.Rows(0)("ExpirationDate"))
            'OrphanTestArgs.GetAppointmentStatus(dtAppt.Rows(i)("EffectiveDate"), dtAppt.Rows(i)("ExpirationDate"), dtAppt.Rows(i)("StatusCode"))

            OrphanTestArgs.GetAppointmentStatus(txtEffectiveDate.Text, txtExpirationDate.Text, ddStatusCode.SelectedValue)



            ' ___ Is this appointment covered by a license?
            ApptIsCovered = LicRules.LicAppt_IsThisApptCoveredByThisLicense(OrphanTestArgs)
            'OurResults.ReportColl.Add("Testing appointment dates against " & State & " license dates. Result: " & IIf(ApptIsCovered, "Pass.", "Fail."))

            ' ___ Adjust appointment effective date to license effective date if license effective date falls after appointment effective date
            If ApptIsCovered Then

                ' ___ Adjust appointment effective date to license effective date if license effective date falls after appointment effective date
                If OrphanTestArgs.LicIsEffective AndAlso OrphanTestArgs.ApptIsActual Then
                    If OrphanTestArgs.LicEffDate > OrphanTestArgs.ApptEffDate Then
                        OrphanTestArgs.ApptEffDate = OrphanTestArgs.LicEffDate
                    End If
                End If

                ' ___ Delete any existing appointments for the enroller in this state for this carrier
                'cCommon.ExecuteNonQuery("DELETE UserAppointments WHERE UserID = '" & cBulkApptSession.UserId & "' AND State = '" & cBulkApptSession.StateTargetList.GetKey(i) & "' AND CarrierID = '" & cBulkApptSession.CarrierId & "'")
                cCommon.ExecuteNonQuery("DELETE UserAppointments WHERE UserID = '" & cBulkApptSession.UserId & "' AND State = '" & State & "' AND CarrierID = '" & CarrierID & "'")

                ' ___ Save record
                '   Public Sub BulkAdd(SaveType, State, CarrierID, Succes,  Comments, Optional MakeFirstRow As Boolean = False)
                PerformSave3(SaveResults, State, CarrierID, OrphanTestArgs)
                If IsDate(txtEffectiveDate.Text) AndAlso cJITAppt.IsJITTrigger(cBulkApptSession.UserId, State, CarrierID) Then
                    cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.ApptExpressJITTrigger, State, CarrierID, 1, Nothing)
                Else
                    cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.ApptExpress, State, CarrierID, 1, Nothing)
                End If

            Else
                cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.ApptExpress, State, CarrierID, 0, "Appt ineligible because of license dates.")
            End If

        Catch ex As Exception
            Throw New Exception("Error #707: BulkAppointment PerformSave2. " & ex.Message)
        End Try
    End Sub

    Private Sub PerformSave3(ByRef SaveResults As LocalResults, ByVal State As String, ByVal CarrierID As String, ByRef OrphanTestArgs As LicRules.OrphanTestArgs)
        Dim Sql As New System.Text.StringBuilder

        Try

            'If OrphanTestArgs.ApptEffDate.ToString("yyyy") = "0001" Then
            '    NewStatusCode = "P"
            'Else
            '    NewStatusCode = "X"
            'End If

            Sql.Append("INSERT INTO UserAppointments (UserID, State, ApplicationDate, EffectiveDate, ExpirationDate, CarrierID, AppointmentNumber, StatusCode, StatusCodeLastChangeDate, AddDate, ChangeDate)")
            Sql.Append(" Values ")
            Sql.Append("(" & cCommon.StrOutHandler(cBulkApptSession.UserId, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(State, False, StringTreatEnum.SideQts_SecApost) & ", ")

            If ddStatusCode.SelectedValue = "X" Then
                Sql.Append("null, ")
            Else
                Sql.Append(cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
            End If

            Sql.Append(cCommon.DateOutHandler(OrphanTestArgs.ApptEffDate, True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(OrphanTestArgs.ApptExpDate, True, True) & ", ")
            Sql.Append(cCommon.StrOutHandler(CarrierID, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(txtAppointmentNumber.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'" & ddStatusCode.SelectedValue & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "') ")

            cCommon.ExecuteNonQuery(Sql.ToString)
            SaveResults.Success = True
            SaveResults.Msg = "Update complete."

        Catch ex As Exception
            SaveResults.Success = False
            SaveResults.Msg = "Unable to save record(s)."
        End Try
    End Sub

    Private Sub DisplayPage(ByVal PageMode As PageMode, ByRef ResponseAction As LocalResponseAction)
        Try

            ' ___ Heading/UserID
            Select Case ResponseAction
                Case LocalResponseAction.DisplayBlank, LocalResponseAction.DisplayUserInputNew
                    litHeading.Text = "Appointment Express"
                Case LocalResponseAction.DisplayExisting, LocalResponseAction.DisplayUserInputExisting
                    If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                        litHeading.Text = "Appointment Express"
                    Else
                        litHeading.Text = "Appointment Express"
                    End If
            End Select

            'HandlePerStatePerCarrier(ResponseAction)

            FormatControls(ResponseAction)

            HandleRecordSelection(PageMode, ResponseAction)

            If ResponseAction = LocalResponseAction.DisplayBlank Then
                HandleData(PageMode, True)
            Else
                HandleData(PageMode, False)
            End If

            HandleStatusAndUpdate()

            ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting

        Catch ex As Exception
            Throw New Exception("Error #709: BulkAppointment DisplayPage. " & ex.Message)
        End Try
    End Sub

    'Private Sub HandlePerStatePerCarrier(ByVal ResponseAction As LocalResponseAction)
    '    If ResponseAction = LocalResponseAction.CarriersPerState Then
    '        cBulkApptSession.PerStatePerCarrier = BulkApptSession.PerStatePerCarrierEnum.CarriersPerState
    '    ElseIf ResponseAction = LocalResponseAction.StatesPerCarrier Then
    '        cBulkApptSession.PerStatePerCarrier = BulkApptSession.PerStatePerCarrierEnum.StatesPerCarrier
    '    End If
    '    Select Case cBulkApptSession.PerStatePerCarrier
    '        Case BulkApptSession.PerStatePerCarrierEnum.CarriersPerState
    '            picCarriersPerState.Visible = True
    '            picStatesPerCarrier.Visible = False
    '            lstCarriersTarget.Visible = True
    '            txtCarrierTarget.Visible = False
    '            lstStatesTarget.Visible = False
    '            txtStateTarget.Visible = True
    '        Case BulkApptSession.PerStatePerCarrierEnum.StatesPerCarrier
    '            picCarriersPerState.Visible = False
    '            picStatesPerCarrier.Visible = True
    '            lstCarriersTarget.Visible = False
    '            txtCarrierTarget.Visible = True
    '            lstStatesTarget.Visible = True
    '            txtStateTarget.Visible = False
    '    End Select
    'End Sub

    Private Sub HandleStatusAndUpdate()
        Try

            txtStatus.Text = String.Empty
            If txtEnrollerTarget.Text.Length = 0 Then
                txtStatus.Text = "no enroller"
            End If
            If cBulkApptSession.CarrierTarget = String.Empty Then
                If txtStatus.Text.Length = 0 Then
                    txtStatus.Text &= "no carrier"
                Else
                    txtStatus.Text &= ", no carrier"
                End If
            End If
            If cBulkApptSession.StateTarget = String.Empty Then
                If txtStatus.Text.Length = 0 Then
                    txtStatus.Text &= "no state"
                Else
                    txtStatus.Text &= ", no state"
                End If
            End If
            If txtStatus.Text.Length > 0 Then
                txtStatus.Text = "N" & txtStatus.Text.Substring(1) & " specified."
            End If

            If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                If txtStatus.Text.Length = 0 Then
                    litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"
                Else
                    litUpdate.Text = "<input onclick='Update()' type='button' value='Update' disabled>"
                End If
            End If
            If litUpdate.Text.Length > 0 Then
                litUpdate.Text &= "&nbsp;&nbsp;<input onclick='GoToLicense()' type='button' value='Go to License'>"
            Else
                litUpdate.Text = "<input onclick='GoToLicense()' type='button' value='Go to License'>"
            End If

        Catch ex As Exception
            Throw New Exception("Error #710: BulkAppointment HandleStatusAndUpdate. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleData(ByVal PageMode As PageMode, ByVal BlankFields As Boolean)
        Try

            If BlankFields Then
                txtAppointmentNumber.Text = String.Empty
                txtApplicationDate.Text = String.Empty
                txtEffectiveDate.Text = String.Empty
                txtExpirationDate.Text = String.Empty
                cBulkApptSession.StatusCode = String.Empty
                cBulkApptSession.AppointmentNumber = String.Empty
                cBulkApptSession.ApplicationDate = #12:00:00 AM#
                cBulkApptSession.EffectiveDate = #12:00:00 AM#
                cBulkApptSession.ExpirationDate = #12:00:00 AM#

                ' ___ StatusCode
                If ddStatusCode.Items.Count = 0 Then
                    ddStatusCode.Items.Add(New ListItem("", ""))
                    ddStatusCode.Items.Add(New ListItem("Pending", "P"))
                    ddStatusCode.Items.Add(New ListItem("Effective", "X"))
                    ddStatusCode.Items.Add(New ListItem("Term", "T"))
                End If

            Else

                Select Case PageMode
                    Case PageMode.Initial
                        cBulkApptSession.Initialized = True

                    Case PageMode.Postback
                        cBulkApptSession.StatusCode = ddStatusCode.SelectedValue
                        cBulkApptSession.AppointmentNumber = txtAppointmentNumber.Text
                        If IsDate(txtApplicationDate.Text) Then
                            cBulkApptSession.ApplicationDate = txtApplicationDate.Text
                        End If
                        If IsDate(txtEffectiveDate.Text) Then
                            cBulkApptSession.EffectiveDate = txtEffectiveDate.Text
                        End If
                        If IsDate(txtExpirationDate.Text) Then
                            cBulkApptSession.ExpirationDate = txtExpirationDate.Text
                        End If

                    Case PageMode.CalledByOther
                        ddStatusCode.SelectedValue = cBulkApptSession.StatusCode
                        txtAppointmentNumber.Text = cBulkApptSession.AppointmentNumber
                        If cBulkApptSession.ApplicationDate <> "12:00:00 AM" Then
                            txtApplicationDate.Text = cBulkApptSession.ApplicationDate
                        End If
                        If cBulkApptSession.EffectiveDate <> "12:00:00 AM" Then
                            txtEffectiveDate.Text = cBulkApptSession.EffectiveDate
                        End If
                        If cBulkApptSession.ExpirationDate <> "12:00:00 AM" Then
                            txtExpirationDate.Text = cBulkApptSession.ExpirationDate
                        End If
                End Select

            End If

        Catch ex As Exception
            Throw New Exception("Error #711: BulkAppointment HandleData. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleRecordSelection(ByVal PageMode As PageMode, ByVal ResponseAction As LocalResponseAction)
        Dim li As ListItem
        Dim DoStateRefresh As Boolean = True

        Try

            Select Case PageMode
                Case PageMode.Initial
                    cBulkApptSession.Letter = "A"
                    'cBulkApptSession.PerStatePerCarrier = BulkApptSession.PerStatePerCarrierEnum.CarriersPerState
                    'lstCarriersTarget.Visible = True
                    'lstStatesTarget.Visible = False
                    cBulkApptSession.CarrierTarget = String.Empty
                    cBulkApptSession.CarrierTargetList.Clear()
                    cBulkApptSession.StateTarget = String.Empty
                    cBulkApptSession.StateTargetList.Clear()

                Case PageMode.Postback

                    Select Case ResponseAction

                        'Case LocalResponseAction.CarriersPerState
                        '        cBulkApptSession.PerStatePerCarrier = BulkApptSession.PerStatePerCarrierEnum.CarriersPerState
                        '        lstCarriersTarget.Visible = True
                        '        lstStatesTarget.Visible = False
                        '        cBulkApptSession.CarrierTarget = String.Empty
                        '        cBulkApptSession.CarrierTargetList.Clear()
                        '        cBulkApptSession.StateTarget = String.Empty
                        '        cBulkApptSession.StateTargetList.Clear()

                        '    Case LocalResponseAction.StatesPerCarrier
                        '        cBulkApptSession.PerStatePerCarrier = BulkApptSession.PerStatePerCarrierEnum.StatesPerCarrier
                        '        lstCarriersTarget.Visible = False
                        '        lstStatesTarget.Visible = True
                        '        cBulkApptSession.CarrierTarget = String.Empty
                        '        cBulkApptSession.CarrierTargetList.Clear()
                        '        cBulkApptSession.StateTarget = String.Empty
                        '        cBulkApptSession.StateTargetList.Clear()

                        Case LocalResponseAction.EnrollerLetterClick
                            cBulkApptSession.Letter = Request.Form("hdSubAction")
                            If cBulkApptSession.Letter = Nothing OrElse cBulkApptSession.Letter = String.Empty Then
                                cBulkApptSession.Letter = "A"
                            End If

                        Case LocalResponseAction.EnrollerAdd
                            If ddEnrollersSource.SelectedIndex > 0 Then
                                cBulkApptSession.UserId = ddEnrollersSource.SelectedItem.Value
                                cBulkApptSession.EnrollerTarget = ddEnrollersSource.SelectedItem.Text
                                cBulkApptSession.CarrierTarget = String.Empty
                                cBulkApptSession.CarrierTargetList.Clear()
                                cBulkApptSession.StateTarget = String.Empty
                                cBulkApptSession.StateTargetList.Clear()
                            End If

                        Case LocalResponseAction.EnrollerRemove
                            cBulkApptSession.UserId = String.Empty
                            cBulkApptSession.EnrollerTarget = String.Empty
                            cBulkApptSession.CarrierTarget = String.Empty
                            cBulkApptSession.CarrierTargetList.Clear()
                            cBulkApptSession.StateTarget = String.Empty
                            cBulkApptSession.StateTargetList.Clear()

                        Case LocalResponseAction.CarrierAdd
                            If lstCarriersSource.SelectedIndex > -1 Then
                                cBulkApptSession.CarrierTarget = lstCarriersSource.SelectedValue
                                cBulkApptSession.CarrierTargetList.Add(lstCarriersSource.SelectedItem.Value, lstCarriersSource.SelectedItem.Text)
                            End If

                        Case LocalResponseAction.CarrierRemove
                            If cBulkApptSession.CarrierTargetList.Count > 0 AndAlso lstCarriersTarget.SelectedIndex > -1 Then
                                For Each li In lstCarriersTarget.Items
                                    If li.Value = lstCarriersTarget.SelectedItem.Value Then
                                        cBulkApptSession.CarrierTargetList.Remove(li.Value)
                                        cBulkApptSession.CarrierTarget = String.Empty
                                        Exit For
                                    End If
                                Next
                            End If

                        Case LocalResponseAction.StateAdd
                            If lstStatesSource.SelectedIndex > -1 Then
                                cBulkApptSession.StateTarget = lstStatesSource.SelectedItem.Value
                                cBulkApptSession.StateTargetList.Add(lstStatesSource.SelectedItem.Value, lstStatesSource.SelectedItem.Text)
                            End If

                        Case LocalResponseAction.StateRemove
                            If cBulkApptSession.StateTargetList.Count > 0 AndAlso lstStatesTarget.SelectedIndex > -1 Then
                                For Each li In lstStatesTarget.Items
                                    If li.Value = lstStatesTarget.SelectedItem.Value Then
                                        cBulkApptSession.StateTargetList.Remove(li.Value)
                                        cBulkApptSession.StateTarget = String.Empty
                                        Exit For
                                    End If
                                Next
                            End If

                        Case LocalResponseAction.StateSourceIndexChange
                            If lstStatesSource.SelectedIndex = -1 Then
                                cBulkApptSession.StateSourceSelectedState = String.Empty
                            Else
                                cBulkApptSession.StateSourceSelectedState = lstStatesSource.SelectedValue
                            End If
                            DoStateRefresh = False

                    End Select

                Case PageMode.CalledByOther
                    'DoStateRefresh = False


            End Select

            RefreshEnroller()
            RefreshCarrier()
            If DoStateRefresh Then
                RefreshState()
            End If

        Catch ex As Exception
            Throw New Exception("Error #712: BulkAppointment HandleRecordSelection. " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshEnroller()
        Dim i As Integer
        Dim li As ListItem
        Dim Sql As String
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable

        Try

            txtEnrollerTarget.Text = cBulkApptSession.EnrollerTarget

            Sql = "SELECT UserId, LastName, FirstName, Role, StatusCode FROM Users "
            Sql &= "WHERE Upper(StatusCode) = 'ACTIVE' AND Upper(Role) IN ('ENROLLER', 'SUPERVISOR') AND left(lastname, 1) = '" & cBulkApptSession.Letter & "' AND UserId <> '" & cBulkApptSession.UserId & "' "
            Sql &= "Order By LastName + FirstName"

            QueryPack = cCommon.GetDTWithQueryPack(Sql, cEnviro.DBHost, cEnviro.DefaultDatabase)
            dt = QueryPack.dt

            ddEnrollersSource.Items.Clear()
            ddEnrollersSource.Items.Add(New ListItem("", 0))
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    li = New ListItem(dt.Rows(i)("LastName") & ", " & dt.Rows(i)("FirstName"), dt.Rows(i)("UserId"))
                    ddEnrollersSource.Items.Add(li)
                Next
            End If

        Catch ex As Exception
            Throw New Exception("Error #713: BulkAppointment RefreshEnroller. " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshCarrier()
        Dim i As Integer
        Dim li As ListItem
        Dim Sql As New System.Text.StringBuilder
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable
        Dim CarrierIdString As String = String.Empty

        Try

            'If cBulkApptSession.UserId = String.Empty Then
            '    cBulkApptSession.StateTarget = String.Empty
            '    txtStateTarget.Text = String.Empty
            '    lstStatesTarget.Items.Clear()
            '    Exit Sub
            'End If

            ' ___ Clear out the controls
            lstCarriersSource.Items.Clear()
            lstCarriersTarget.Items.Clear()

            ' ___ Populate the target controls
            With lstCarriersTarget
                .DataSource = cBulkApptSession.CarrierTargetList
                .DataValueField = "Key"    ' Value
                .DataTextField = "Key"  ' Text
                ' .DataTextField = "Value"  ' Text
                .DataBind()
            End With

            ' ___ Get a list of selected carriers to exclude from the source carrier list
            If cBulkApptSession.CarrierTargetList.Count > 0 Then
                For i = 0 To cBulkApptSession.CarrierTargetList.Count - 1
                    CarrierIdString &= "'" & cBulkApptSession.CarrierTargetList.GetKey(i) & "'"
                    If i < cBulkApptSession.CarrierTargetList.Count - 1 Then
                        CarrierIdString &= ", "
                    End If
                Next
            End If

            ' ___ Build the sql
            If cBulkApptSession.StateSourceSelectedState = String.Empty Then
                Sql.Append("SELECT CarrierID FROM Codes_CarrierID WHERE CarrierId Not In ('" & cBulkApptSession.CarrierTarget & "')  ORDER BY CarrierID ")
            Else
                Sql.Append("SELECT cc.CarrierID, ")
                Sql.Append("Status = case ")
                Sql.Append("when ua.CarrierID is null then 'None' ")
                Sql.Append("when ua.EffectiveDate is null then 'Pending' ")
                Sql.Append("else 'Effective' ")
                Sql.Append("End ")
                Sql.Append("FROM Codes_CarrierID cc ")
                Sql.Append("LEFT JOIN UserAppointments ua on cc.CarrierID = ua.CarrierID AND ua.UserID = '" & cBulkApptSession.UserId & "' AND ua.State = '" & lstStatesSource.SelectedValue & "' ")
                If CarrierIdString.Length > 0 Then
                    Sql.Append("WHERE cc.CarrierId Not In (" & CarrierIdString & ") ")
                End If
                Sql.Append("ORDER BY cc.CarrierID")
            End If

            ' ___ Execute the query
            QueryPack = cCommon.GetDTWithQueryPack(Sql.ToString, cEnviro.DBHost, cEnviro.DefaultDatabase)
            dt = QueryPack.dt

            If dt.Rows.Count > 0 Then
                If lstStatesSource.SelectedIndex = -1 Then
                    For i = 0 To dt.Rows.Count - 1
                        li = New ListItem(dt.Rows(i)("CarrierId"), dt.Rows(i)("CarrierId"))
                        lstCarriersSource.Items.Add(li)
                    Next
                Else
                    For i = 0 To dt.Rows.Count - 1
                        li = New ListItem(dt.Rows(i)("CarrierId") & " - " & dt.Rows(i)("Status"), dt.Rows(i)("CarrierId"))
                        lstCarriersSource.Items.Add(li)
                    Next
                End If
            End If

            '' ___ Populate the source control
            'With lstCarriersSource
            '    .DataSource = dt
            '    .DataValueField = "CarrierID"    ' Value
            '    .DataTextField = "Status"  ' Text
            '    .DataBind()
            'End With

        Catch ex As Exception
            Throw New Exception("Error #714: BulkAppointment RefreshCarrier. " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshState()
        Dim i As Integer
        Dim li As ListItem
        Dim Sql As New System.Text.StringBuilder
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable
        Dim StateIdString As String = String.Empty
        Dim InnerJoin As String = String.Empty

        Try

            ' ___ Clear out the controls
            lstStatesSource.Items.Clear()
            lstStatesTarget.Items.Clear()

            ' ___ Populate the target controls
            With lstStatesTarget
                .DataSource = cBulkApptSession.StateTargetList
                .DataValueField = "Key"
                .DataTextField = "Value"
                .DataBind()
            End With

            ' ___ Get a list of selected states to exclude from the source states list
            If cBulkApptSession.StateTargetList.Count > 0 Then
                For i = 0 To cBulkApptSession.StateTargetList.Count - 1
                    StateIdString &= "'" & cBulkApptSession.StateTargetList.GetKey(i) & "'"
                    If i < cBulkApptSession.StateTargetList.Count - 1 Then
                        StateIdString &= ", "
                    End If
                Next
            End If

            ' ___ Build the sql
            InnerJoin = " INNER JOIN UserLicenses ul ON ul.UserId = '" & cBulkApptSession.UserId & "' AND  ul.State = cs.StateCode "
            Sql.Append("SELECT cs.StateCode, ")
            Sql.Append("Status = case ")
            Sql.Append("when ul.EffectiveDate = '1/1/1950' then 'Pending' ")
            Sql.Append("else 'Effective' ")
            Sql.Append("end ")
            Sql.Append("FROM Codes_State cs" & InnerJoin & " ")
            'Sql.Append("WHERE cs.LogicalDelete = 0  ")
            Sql.Append("WHERE ul.ExpirationDate IS NULL or ul.ExpirationDate >= '" & cCommon.GetServerDateTime & "' ")
            If lstStatesTarget.Items.Count > 0 Then
                Sql.Append("AND StateCode Not In (" & StateIdString & ") ")
            End If
            Sql.Append("Order By cs.StateCode")

            ' ___ Execute the query
            QueryPack = cCommon.GetDTWithQueryPack(Sql.ToString, cEnviro.DBHost, cEnviro.DefaultDatabase)
            dt = QueryPack.dt

            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    li = New ListItem(dt.Rows(i)("StateCode") & " - " & dt.Rows(i)("Status"), dt.Rows(i)("StateCode"))
                    lstStatesSource.Items.Add(li)
                Next
            End If

            '' ___ Populate the source control
            'With lstStatesSource
            '    .DataSource = dt
            '    .DataValueField = "StateCode"
            '    .DataTextField = "StateCode"
            '    .DataBind()
            'End With

        Catch ex As Exception
            Throw New Exception("Error #715: BulkAppointment RefreshState. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls(ByVal ResponseAction As LocalResponseAction)
        Try

            lstStatesSource.Attributes.Add("OnChange", "StateSourceIndexChange()")

            ddEnrollersSource.Attributes.Add("style", "width:230")
            lstCarriersSource.Attributes.Add("style", "width:230")
            lstCarriersTarget.Attributes.Add("style", "width:230")
            lstStatesSource.Attributes.Add("style", "width:230")
            lstStatesTarget.Attributes.Add("style", "width:230")

            btnEnrollerAdd.Attributes.Add("onclick", "ChangeSelection('enrolleradd')")
            btnEnrollerRemove.Attributes.Add("onclick", "ChangeSelection('enrollerremove')")
            btnCarrierAdd.Attributes.Add("onclick", "ChangeSelection('carrieradd')")
            btnCarrierRemove.Attributes.Add("onclick", "ChangeSelection('carrierremove')")
            btnStateAdd.Attributes.Add("onclick", "ChangeSelection('stateadd')")
            btnStateRemove.Attributes.Add("onclick", "ChangeSelection('stateremove')")

            Style.AddStyle(txtEnrollerTarget, Style.StyleType.NoneditableWhite, 230)
            'Style.AddStyle(txtCarrierTarget, Style.StyleType.NoneditableWhite, 230)

            Style.AddStyle(txtStatus, Style.StyleType.NoneditableGrayed, 300)

            ' ___ Edit
            If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"

                lblApplicationDateLink.Visible = True
                lblEffectiveDateLink.Visible = True
                lblExpirationDateLink.Visible = True

                Style.AddStyle(txtAppointmentNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableWhite, 300)

                btnEnrollerAdd.Enabled = True
                btnEnrollerRemove.Enabled = True
                btnCarrierAdd.Enabled = True
                btnCarrierRemove.Enabled = True
                btnStateAdd.Enabled = True
                btnStateRemove.Enabled = True

            Else

                lblApplicationDateLink.Visible = False
                lblEffectiveDateLink.Visible = False
                lblExpirationDateLink.Visible = False

                Style.AddStyle(txtAppointmentNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableGrayed, 300)

                btnEnrollerAdd.Enabled = False
                btnEnrollerRemove.Enabled = False
                btnCarrierAdd.Enabled = False
                btnCarrierRemove.Enabled = False
                btnStateAdd.Enabled = False
                btnStateRemove.Enabled = False
            End If

        Catch ex As Exception
            Throw New Exception("Error #716: BulkAppointment FormatControls. " & ex.Message)
        End Try
    End Sub

    Private Class LocalResults
        Private cSuccess As Boolean
        Private cMsg As String
        Private cResponseAction As LocalResponseAction
        Private cObtainConfirm As Boolean
        Public Property Success() As Boolean
            Get
                Return cSuccess
            End Get
            Set(ByVal Value As Boolean)
                cSuccess = Value
            End Set
        End Property
        Public Property Msg() As String
            Get
                Return cMsg
            End Get
            Set(ByVal Value As String)
                cMsg = Value
            End Set
        End Property
        Public Property ResponseAction() As LocalResponseAction
            Get
                Return cResponseAction
            End Get
            Set(ByVal Value As LocalResponseAction)
                cResponseAction = Value
            End Set
        End Property
        Public Property ObtainConfirm() As Boolean
            Get
                Return cObtainConfirm
            End Get
            Set(ByVal Value As Boolean)
                cObtainConfirm = Value
            End Set
        End Property
    End Class
End Class


'Public Class DateTimeBVI
'    Private cValue As Object = DBNull.Value
'    Public Property Value()
'        Get
'            If IsDBNull(cValue) Then
'                Return DBNull.Value
'            Else
'                Return CType(cValue, DateTime)
'            End If
'        End Get
'        Set(ByVal Value As Object)
'            If IsDBNull(Value) Then
'                cValue = DBNull.Value
'            Else
'                cValue = CType(Value, DateTime)
'            End If
'        End Set
'    End Property
'End Class
