Imports System.Data
Imports System.Data.SqlClient

Partial Class BulkLicense
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected WithEvents litMsg As System.Web.UI.WebControls.Literal
    Protected WithEvents litResponseAction As System.Web.UI.WebControls.Literal
    Protected WithEvents lblCurrentRights As System.Web.UI.WebControls.Label
    Protected WithEvents lblState As System.Web.UI.WebControls.Label
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Protected WithEvents litEnviro As System.Web.UI.WebControls.Literal
    Private cBulkLicSession As BulkLicSession
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack
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
        StateAdd = 14
        StateRemove = 15
        GoToAppt = 16
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
        StateAdd = 10
        StateRemove = 11
        GoToAppt = 12
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

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(RightsClass.LicenseEdit, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cBulkLicSession = Session("BulkLicSession")

            ' ___ Initialize the BulkAppointmentDataPack (used to display JIT results)
            cBulkAppointmentDataPack = Session("BulkAppointmentDataPack")
            cBulkAppointmentDataPack.Init()

            ' ___ Get the page mode
            PageMode = cCommon.GetPageMode(Page, cBulkLicSession)

            ' ___ Get RequestAction
            RequestAction = GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(PageMode, RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = LocalResponseAction.ReturnToCallingPage Then
                'cSess.PageReturnOnLoadMessage = Results.Msg
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

                'If Not Results.Msg = Nothing Then
                '    litMsg.Text = "<script type=""text/javascript"">alert('" & Results.Msg & "')</script>"
                'End If


                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.Text = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            Dim ErrorObj As New ErrorObj(ex, "Error #602: BulkLicense Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function GetRequestAction(ByVal Page As Page) As LocalRequestAction
        Dim ActionType As String = Nothing
        Dim RequestAction As LocalRequestAction
        Dim hdResponseAction As String = Nothing
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
                ElseIf hdAction = "stateadd" Then
                    ActionType = "selection"
                ElseIf hdAction = "stateremove" Then
                    ActionType = "selection"
                ElseIf hdAction = "gotoappt" Then
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
                        Case "stateadd"
                            RequestAction = LocalRequestAction.StateAdd
                        Case "stateremove"
                            RequestAction = LocalRequestAction.StateRemove
                        Case "gotoappt"
                            RequestAction = LocalRequestAction.GoToAppt
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
                    hdResponseAction = hdResponseAction = Page.Request.Form("hdResponseAction")
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
            Throw New Exception("Error #603: BulkLicense GetRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function ExecuteRequestAction(ByVal PageMode As PageMode, ByVal RequestAction As LocalRequestAction) As LocalResults
        Dim ValidationResults As New LocalResults
        Dim SaveResults As New LocalResults
        Dim MyResults As New LocalResults

        Try

            Select Case RequestAction
                Case LocalRequestAction.ReturnToParent
                    MyResults.ResponseAction = LocalResponseAction.ReturnToCallingPage
                    MyResults.Success = True

                Case LocalRequestAction.CreateNew
                    MyResults.ResponseAction = LocalResponseAction.DisplayBlank

                    'Case RequestAction.SaveNew
                    '    ValidationResults = IsDataValid(RequestAction)
                    '    If ValidationResults.Success Then
                    '        SaveResults = PerformSave(RequestAction)
                    '        If SaveResults.Success Then
                    '            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    '            MyResults.Msg = SaveResults.Msg
                    '            ' MyResults.Msg = GetBulkResults(SaveResults)
                    '        Else
                    '            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    '            MyResults.Msg = SaveResults.Msg
                    '        End If
                    '    ElseIf ValidationResults.ObtainConfirm Then
                    '        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    '        MyResults.Msg = ValidationResults.Msg
                    '    Else
                    '        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    '        MyResults.Msg = ValidationResults.Msg
                    '    End If

                    'Case RequestAction.LoadExisting
                    '    MyResults.ResponseAction = LocalResponseAction.DisplayExisting

                    'Case RequestAction.SaveExisting
                    '    ValidationResults = IsDataValid(RequestAction)
                    '    If ValidationResults.Success Then
                    '        SaveResults = PerformSave(RequestAction)
                    '        If SaveResults.Success Then
                    '            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNew
                    '            MyResults.Msg = SaveResults.Msg
                    '            'MyResults.Msg = GetBulkResults(SaveResults)
                    '        Else
                    '            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                    '            MyResults.Msg = SaveResults.Msg
                    '        End If
                    '    ElseIf ValidationResults.ObtainConfirm Then
                    '        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                    '        MyResults.Msg = ValidationResults.Msg
                    '    Else
                    '        MyResults.ResponseAction = LocalResponseAction.DisplayExisting
                    '        MyResults.Msg = ValidationResults.Msg
                    '    End If

                Case LocalRequestAction.SaveNewOrExisting
                    ValidationResults = IsDataValid(RequestAction)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting
                            MyResults.Msg = SaveResults.Msg
                        Else
                            MyResults.ResponseAction = LocalResponseAction.DisplayUserInputExisting
                            MyResults.Msg = SaveResults.Msg
                        End If
                    ElseIf ValidationResults.ObtainConfirm Then
                        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting
                        MyResults.ObtainConfirm = True
                        MyResults.Msg = ValidationResults.Msg
                    Else
                        MyResults.ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

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

                Case LocalRequestAction.StateAdd
                    MyResults.ResponseAction = LocalResponseAction.StateAdd

                Case LocalRequestAction.StateRemove
                    MyResults.ResponseAction = LocalResponseAction.StateRemove

                Case LocalRequestAction.GoToAppt
                    HandleData(PageMode)
                    Response.Redirect("BulkAppointment.aspx?CalledBy=Other")

                Case LocalRequestAction.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #604: BulkLicense ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function GetBulkResults(ByRef SaveResults As BulkResults) As String
        Dim i As Integer
        Dim Msg As String = String.Empty
        With SaveResults.ReportColl
            For i = 1 To .Count
                Msg &= .Item(i) & "\n"
            Next
        End With
        Return Msg
    End Function

    Private Function IsDataValid(ByVal RequestAction As LocalRequestAction) As LocalResults
        Dim MyResults As New LocalResults
        Dim ErrColl As New Collection
        Dim OKToProceed As Boolean = True
        Dim LicRules As New LicRules(cCommon)

        Try

            ' ___ Trim the textbox input
            txtLicenseNumber.Text = Trim(txtLicenseNumber.Text)
            txtNotes.Text = Trim(txtNotes.Text)

            'cCommon.ValidateDropDown(ErrColl, ddState, 1, "state not provided")

            If txtEffectiveDate.Text.Length > 0 And txtLicenseNumber.Text.Length = 0 Then
                cCommon.ValidateErrorOnly(ErrColl, "a license must have a license number")
            End If

            LicRules.LicAppt_LongTermCareStateSpecificDateCheck(ErrColl, cBulkLicSession.StateTarget, txtLongTermCareStateSpecificEffectiveDate, txtLongTermCareStateSpecificExpirationDate)
            LicRules.LicAppt_DateSequenceCheck(ErrColl, txtEffectiveDate, txtExpirationDate)
            LicRules.Lic_KeyViolationCheck(RequestAction, ErrColl, cBulkLicSession.UserId, cBulkLicSession.StateTarget)
            LicRules.Lic_AllowThisEdit(ErrColl, cBulkLicSession.UserId, cBulkLicSession.StateTarget, txtEffectiveDate.Text, txtExpirationDate.Text)

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #605: BulkLicense IsDataValid. " & ex.Message)
        End Try
    End Function

    'Private Function DeleteLicense()
    '    Dim MyResults As New Results
    '    Dim Querypack As DBase.QueryPack
    '    Dim Sql As String

    '    Sql = "DELETE UserLicenses WHERE UserID = '" & cBulkLicSession.UserId & "' AND State = '" & cBulkLicSession.StateTarget & "'"
    '    Querypack = cCommon.ExecuteNonQueryWithQuerypack(cEnviro.DBHost, cEnviro.DefaultDatabase, Sql)
    '    MyResults.Success = Querypack.Success
    '    If Querypack.Success Then
    '        MyResults.Msg = "Record deleted."
    '    Else
    '        MyResults.Msg = "Unable to delete record."
    '    End If
    '    Return MyResults

    'End Function

    Private Function DeleteLicense() As LocalResults
        Dim MyResults As New LocalResults
        Dim QueryPack As DBase.QueryPack
        Dim Sql As String

        Try

            Sql = "exec usp_DeleteLicense @UserID = '" & cBulkLicSession.UserId & "', @State = '" & cBulkLicSession.StateTarget & "'"
            QueryPack = cCommon.GetDTWithQueryPack(Sql, False)

            MyResults.Success = QueryPack.Success
            If QueryPack.Success Then
                MyResults.Msg = "Record deleted."
            Else
                MyResults.Msg = "Unable to delete record." & QueryPack.TechErrMsg
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #606: BulkLicense DeleteLicense. " & ex.Message)
        End Try
    End Function


    Private Function PerformSave(ByRef RequestAction As RequestActionEnum) As LocalResults
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New LocalResults
        Dim DeleteLicenseResults As LocalResults
        Dim NewLicense As Boolean
        Dim dt As DataTable
        Dim JITAppt As New JITAppt(Me)

        Try

            dt = cCommon.GetDT("SELECT UserID FROM UserManagement..UserLicenses WHERE UserID = '" & cBulkLicSession.UserId & "' AND State = '" & cBulkLicSession.StateTarget & "'")
            If dt.Rows.Count = 0 Then
                NewLicense = True
            End If

            ' ___ First, delete any existing license record.
            DeleteLicenseResults = DeleteLicense()
            If Not DeleteLicenseResults.Success Then
                MyResults.Success = False
                MyResults.Msg = "Error occurred in attempt to delete license for " & cBulkLicSession.UserId & "(" & cBulkLicSession.StateTarget & ")"
                Return MyResults
            End If

            Sql.Append("INSERT INTO UserLicenses (UserID, State, EffectiveDate, ExpirationDate, LongTermCareStateSpecificEffectiveDate, LongTermCareStateSpecificExpirationDate, LicenseNumber, ApplicationDate, RenewalDateSent, RenewalDateRecd, Notes, AddDate, ChangeDate)")
            Sql.Append(" Values ")
            Sql.Append("(" & cCommon.StrOutHandler(cBulkLicSession.UserId, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(cBulkLicSession.StateTarget, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.DateOutHandler(txtEffectiveDate.Text, False, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")

            Sql.Append(cCommon.DateOutHandler(txtLongTermCareStateSpecificEffectiveDate.Text, True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(txtLongTermCareStateSpecificExpirationDate.Text, True, True) & ", ")
            Sql.Append(cCommon.StrOutHandler(txtLicenseNumber.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")

            If IsDate(txtEffectiveDate.Text) Then
                Sql.Append("null, ")
            Else
                Sql.Append(cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
            End If

            Sql.Append(cCommon.DateOutHandler(txtRenewalDateSent.Text, True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(txtRenewalDateRecd.Text, True, True) & ", ")
            Sql.Append(cCommon.StrOutHandler(txtNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ",")
            Sql.Append("'" & cCommon.GetServerDateTime & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "') ")

            Try
                cCommon.ExecuteNonQuery(Sql.ToString)
                MyResults.Success = True
                MyResults.Msg = "Update complete."

                ' ___ Handle ExPostFacto JIT
                If NewLicense AndAlso IsDate(txtEffectiveDate.Text) Then
                    JITAppt.PerformExPostFactoJustInTimeAppointments(cBulkLicSession.UserId, cBulkLicSession.StateTarget, txtEffectiveDate.Text, txtExpirationDate.Text)
                End If

                ' ___ Handle time-sensitive data
                cCommon.UpdateTimeSensitiveData(cBulkLicSession.UserId, cBulkLicSession.StateTarget)

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Msg = "Unable to save record."
            End Try

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #607: BulkLicense PerformSave. " & ex.Message)
        End Try
    End Function


    Private Sub DisplayPage(ByVal PageMode As PageMode, ByRef ResponseAction As LocalResponseAction)
        Dim dt As DataTable = Nothing
        Dim RefreshFromData As Boolean
        Dim DoClear As Boolean

        Try

            ' ___ Set the attributes for controls
            txtNotes.Attributes.Add("onkeyup", "return legallength(this, 512)")

            ' ___ Heading/UserID
            litHeading.Text = "License Express"

            ' FormatControls()

            HandleRecordSelection(PageMode, ResponseAction)

            HandleData(PageMode)

            HandleStatusAndUpdate()

            If Not ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting Then
                If cBulkLicSession.EnrollerTarget <> String.Empty AndAlso cBulkLicSession.StateTarget <> String.Empty Then
                    dt = cCommon.GetDT("SELECT * From UserLicenses Where UserID = '" & cBulkLicSession.UserId & "' AND State = '" & cBulkLicSession.StateTarget & "'")
                    If dt.Rows.Count > 0 Then
                        RefreshFromData = True
                    Else
                        DoClear = True
                    End If
                Else
                    DoClear = True
                End If
            End If


            'dt.Rows(0)("ExpirationDate")

            If RefreshFromData Then

                'If txtStateFullTarget.Text.Length > 0 Then
                '    If dt.Rows(0)("EffectiveDate") = "1/1/1950" Then
                '        cBulkLicSession.StateFullTarget = cBulkLicSession.StateTarget & "- Pending"
                '    Else
                '        cBulkLicSession.StateFullTarget = cBulkLicSession.StateTarget & "- Effective"
                '    End If
                'End If

                If txtStateFullTarget.Text.Length > 0 Then
                    If dt.Rows(0)("EffectiveDate") = "1/1/1950" Then
                        cBulkLicSession.StateFullTarget = cBulkLicSession.StateTarget & "- Pending"
                    ElseIf (Not IsDBNull(dt.Rows(0)("ExpirationDate"))) AndAlso dt.Rows(0)("ExpirationDate") <= cCommon.GetServerDateTime Then
                        cBulkLicSession.StateFullTarget = cBulkLicSession.StateTarget & "- Expired"
                    Else
                        cBulkLicSession.StateFullTarget = cBulkLicSession.StateTarget & "- Effective"
                    End If
                End If

                txtStateFullTarget.Text = cBulkLicSession.StateFullTarget

                txtLicenseNumber.Text = cCommon.StrInHandler(dt.Rows(0)("LicenseNumber"))
                txtApplicationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ApplicationDate"))
                txtEffectiveDate.Text = cCommon.DateInHandler(dt.Rows(0)("EffectiveDate"))
                txtExpirationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ExpirationDate"))
                txtLongTermCareStateSpecificEffectiveDate.Text = cCommon.DateInHandler(dt.Rows(0)("LongTermCareStateSpecificEffectiveDate"))
                txtLongTermCareStateSpecificExpirationDate.Text = cCommon.DateInHandler(dt.Rows(0)("LongTermCareStateSpecificExpirationDate"))
                txtRenewalDateSent.Text = cCommon.DateInHandler(dt.Rows(0)("RenewalDateSent"))
                txtRenewalDateRecd.Text = cCommon.DateInHandler(dt.Rows(0)("RenewalDateRecd"))
                txtNotes.Text = cCommon.StrInHandler(dt.Rows(0)("Notes"))

            ElseIf DoClear Then
                txtLicenseNumber.Text = String.Empty
                txtApplicationDate.Text = String.Empty
                txtEffectiveDate.Text = String.Empty
                txtExpirationDate.Text = String.Empty
                txtLongTermCareStateSpecificEffectiveDate.Text = String.Empty
                txtLongTermCareStateSpecificExpirationDate.Text = String.Empty
                txtRenewalDateSent.Text = String.Empty
                txtRenewalDateRecd.Text = String.Empty
                txtNotes.Text = String.Empty
            End If

            FormatControls()

            ResponseAction = LocalResponseAction.DisplayUserInputNewOrExisting

        Catch ex As Exception
            Throw New Exception("Error #608: BulkLicense DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleRecordSelection(ByVal PageMode As PageMode, ByVal ResponseAction As LocalResponseAction)
        Dim DoStateRefresh As Boolean = True

        Try

            Select Case PageMode
                Case PageMode.Initial
                    cBulkLicSession.Letter = "A"

                Case PageMode.Postback

                    Select Case ResponseAction
                        Case LocalResponseAction.EnrollerLetterClick
                            cBulkLicSession.Letter = Request.Form("hdSubAction")
                            If cBulkLicSession.Letter = Nothing OrElse cBulkLicSession.Letter = String.Empty Then
                                cBulkLicSession.Letter = "A"
                            End If

                        Case LocalResponseAction.EnrollerAdd
                            If ddEnrollersSource.SelectedIndex > 0 Then
                                cBulkLicSession.UserId = ddEnrollersSource.SelectedItem.Value
                                cBulkLicSession.EnrollerTarget = ddEnrollersSource.SelectedItem.Text
                                cBulkLicSession.StateTarget = String.Empty
                                cBulkLicSession.StateFullTarget = String.Empty
                            End If

                        Case LocalResponseAction.EnrollerRemove
                            cBulkLicSession.UserId = String.Empty
                            cBulkLicSession.EnrollerTarget = String.Empty
                            cBulkLicSession.StateTarget = String.Empty
                            cBulkLicSession.StateFullTarget = String.Empty

                        Case LocalResponseAction.StateAdd
                            If lstStatesSource.SelectedIndex > -1 Then
                                cBulkLicSession.StateFullTarget = lstStatesSource.SelectedItem.Text
                                cBulkLicSession.StateTarget = lstStatesSource.SelectedItem.Value
                            End If

                        Case LocalResponseAction.StateRemove
                            cBulkLicSession.StateFullTarget = String.Empty
                            cBulkLicSession.StateTarget = String.Empty

                    End Select

                Case PageMode.CalledByOther
                    ' DoStateRefresh = False

            End Select

            RefreshEnroller()
            If DoStateRefresh Then
                RefreshState()
            End If

        Catch ex As Exception
            Throw New Exception("Error #609: BulkLicense HandleRecordSelection. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleData(ByVal PageMode As PageMode)
        Try

            Select Case PageMode
                Case PageMode.Initial
                    cBulkLicSession.Initialized = True

                Case PageMode.Postback
                    cBulkLicSession.LicenseNumber = txtLicenseNumber.Text
                    If IsDate(txtApplicationDate.Text) Then
                        cBulkLicSession.ApplicationDate = txtApplicationDate.Text
                    End If
                    If IsDate(txtEffectiveDate.Text) Then
                        cBulkLicSession.EffectiveDate = txtEffectiveDate.Text
                    End If
                    If IsDate(txtExpirationDate.Text) Then
                        cBulkLicSession.ExpirationDate = txtExpirationDate.Text
                    End If
                    If IsDate(txtRenewalDateSent.Text) Then
                        cBulkLicSession.RenewalDateSent = txtRenewalDateSent.Text
                    End If
                    If IsDate(txtRenewalDateRecd.Text) Then
                        cBulkLicSession.RenewalDateRecd = txtRenewalDateRecd.Text
                    End If
                    If IsDate(txtLongTermCareStateSpecificEffectiveDate.Text) Then
                        cBulkLicSession.LongTermCareStateSpecificEffectiveDate = txtLongTermCareStateSpecificEffectiveDate.Text
                    End If
                    If IsDate(txtLongTermCareStateSpecificExpirationDate.Text) Then
                        cBulkLicSession.LongTermCareStateSpecificExpirationDate = txtLongTermCareStateSpecificExpirationDate.Text
                    End If
                    cBulkLicSession.Notes = txtNotes.Text

                Case PageMode.CalledByOther
                    cBulkLicSession.Initialized = True

                    txtLicenseNumber.Text = cBulkLicSession.LicenseNumber
                    If cBulkLicSession.ApplicationDate <> "12:00:00 AM" Then
                        txtApplicationDate.Text = cBulkLicSession.ApplicationDate
                    End If
                    If cBulkLicSession.EffectiveDate <> "12:00:00 AM" Then
                        txtEffectiveDate.Text = cBulkLicSession.EffectiveDate
                    End If
                    If cBulkLicSession.ExpirationDate <> "12:00:00 AM" Then
                        txtExpirationDate.Text = cBulkLicSession.ExpirationDate
                    End If
                    If cBulkLicSession.RenewalDateSent <> "12:00:00 AM" Then
                        txtRenewalDateSent.Text = cBulkLicSession.RenewalDateSent
                    End If
                    If cBulkLicSession.RenewalDateRecd <> "12:00:00 AM" Then
                        txtRenewalDateRecd.Text = cBulkLicSession.RenewalDateRecd
                    End If
                    If cBulkLicSession.LongTermCareStateSpecificEffectiveDate <> "12:00:00 AM" Then
                        txtLongTermCareStateSpecificEffectiveDate.Text = cBulkLicSession.LongTermCareStateSpecificEffectiveDate
                    End If
                    If cBulkLicSession.LongTermCareStateSpecificExpirationDate <> "12:00:00 AM" Then
                        txtLongTermCareStateSpecificExpirationDate.Text = cBulkLicSession.LongTermCareStateSpecificExpirationDate
                    End If
                    txtNotes.Text = cBulkLicSession.Notes
            End Select

        Catch ex As Exception
            Throw New Exception("Error #610: BulkLicense HandleData. " & ex.Message)
        End Try
    End Sub

    Private Sub HandleStatusAndUpdate()
        Try

            txtStatus.Text = String.Empty
            If txtEnrollerTarget.Text.Length = 0 Then
                txtStatus.Text = "no enroller"
            End If
            If txtStateFullTarget.Text.Length = 0 Then
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
                litUpdate.Text &= "&nbsp;&nbsp;<input onclick='GoToAppt()' type='button' value='Go to Appt'>"
            Else
                litUpdate.Text = "<input onclick='GoToAppt()' type='button' value='Go to Appt'>"
            End If

        Catch ex As Exception
            Throw New Exception("Error #611: BulkLicense HandleStatusAndUpdate. " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshEnroller()
        Dim i As Integer
        Dim li As ListItem
        Dim Sql As String
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable

        Try

            txtEnrollerTarget.Text = cBulkLicSession.EnrollerTarget

            'Sql = "SELECT UserId, LastName, FirstName, Role, StatusCode From Users WHERE Upper(StatusCode) = 'ACTIVE' AND Upper(Role) = 'ENROLLER' AND left(lastname, 1) = '" & cBulkLicSession.Letter & "' AND UserId <> '" & cBulkLicSession.UserId & "' Order By LastName + FirstName"
            Sql = "SELECT UserId, LastName, FirstName, Role, StatusCode From Users WHERE Upper(StatusCode) = 'ACTIVE' AND Upper(Role) IN ('ENROLLER', 'SUPERVISOR') AND left(lastname, 1) = '" & cBulkLicSession.Letter & "' AND UserId <> '" & cBulkLicSession.UserId & "' Order By LastName + FirstName"
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
            Throw New Exception("Error #612: BulkLicense RefreshEnroller. " & ex.Message)
        End Try
    End Sub

    Private Sub RefreshState()
        Dim i As Integer
        Dim li As ListItem
        Dim Sql As New System.Text.StringBuilder
        Dim QueryPack As DBase.QueryPack
        Dim dt As DataTable

        Try

            'If cBulkLicSession.UserId = String.Empty Then
            '    cBulkLicSession.StateTarget = String.Empty
            '    cBulkLicSession.StateFullTarget = String.Empty
            '    txtStateFullTarget.Text = String.Empty
            '    Exit Sub
            'End If

            txtStateFullTarget.Text = cBulkLicSession.StateFullTarget

            Sql.Append("SELECT cs.StateCode, ul.State, ")
            Sql.Append("Status = case ")
            Sql.Append("when ul.State is null then 'Unlicensed' ")
            Sql.Append("when ul.EffectiveDate = '1/1/1950' then 'Pending' ")
            Sql.Append("when ul.ExpirationDate IS NOT NULL AND ul.ExpirationDate <= '" & cCommon.GetServerDateTime & "' then 'Expired' ")
            Sql.Append("else 'Effective' ")
            Sql.Append("End ")
            Sql.Append("FROM Codes_State cs ")
            Sql.Append("LEFT JOIN UserLicenses ul on cs.StateCode = ul.State AND ul.UserID = '" & cBulkLicSession.UserId & "'")
            'Sql.Append("WHERE cs.LogicalDelete = 0")
            If txtStateFullTarget.Text <> String.Empty Then
                'Sql.Append(" AND cs.StateCode <> '" & cBulkLicSession.StateTarget & "'")
                Sql.Append(" WHERE cs.StateCode <> '" & cBulkLicSession.StateTarget & "'")
            End If

            QueryPack = cCommon.GetDTWithQueryPack(Sql.ToString, cEnviro.DBHost, cEnviro.DefaultDatabase)
            dt = QueryPack.dt

            lstStatesSource.Items.Clear()
            If dt.Rows.Count > 0 Then
                For i = 0 To dt.Rows.Count - 1
                    li = New ListItem(dt.Rows(i)("StateCode") & " - " & dt.Rows(i)("Status"), dt.Rows(i)("StateCode"))
                    lstStatesSource.Items.Add(li)
                Next
            End If

        Catch ex As Exception
            Throw New Exception("Error #613: BulkLicense RefreshState. " & ex.Message)
        End Try
    End Sub

    Private Sub FormatControls()
        Try

            ddEnrollersSource.Attributes.Add("style", "width:230")
            lstStatesSource.Attributes.Add("style", "width:230")
            Style.AddStyle(txtStateFullTarget, Style.StyleType.NoneditableWhite, 230)
            btnEnrollerAdd.Attributes.Add("onclick", "ChangeSelection('enrolleradd')")
            btnEnrollerRemove.Attributes.Add("onclick", "ChangeSelection('enrollerremove')")
            btnStateAdd.Attributes.Add("onclick", "ChangeSelection('stateadd')")
            btnStateRemove.Attributes.Add("onclick", "ChangeSelection('stateremove')")
            Style.AddStyle(txtEnrollerTarget, Style.StyleType.NoneditableWhite, 230)
            Style.AddStyle(txtStatus, Style.StyleType.NoneditableGrayed, 300)

            If cCommon.StateRequiresStateSpecificLTCCert(cBulkLicSession.StateTarget) Then
                plLongTermCareStateSpecific.Visible = True
            Else
                plLongTermCareStateSpecific.Visible = False
            End If

            ' ___ Edit
            If cRights.HasThisRight(RightsClass.LicenseEdit) Then

                lblApplicationDateLink.Visible = True
                lblEffectiveDateLink.Visible = True
                lblExpirationDateLink.Visible = True
                lblRenewalDateSentLink.Visible = True
                lblRenewalDateRecdLink.Visible = True
                lblLongTermCareStateSpecificEffectiveDateLink.Visible = True
                lblLongTermCareStateSpecificExpirationDateLink.Visible = True

                Style.AddStyle(txtLicenseNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtRenewalDateSent, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtRenewalDateRecd, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtLongTermCareStateSpecificEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtLongTermCareStateSpecificExpirationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtNotes, Style.StyleType.NormalEditable, 400, True)

                btnEnrollerAdd.Enabled = True
                btnEnrollerRemove.Enabled = True
                btnStateAdd.Enabled = True
                btnStateRemove.Enabled = True

            Else

                lblApplicationDateLink.Visible = False
                lblEffectiveDateLink.Visible = False
                lblExpirationDateLink.Visible = False
                lblRenewalDateSentLink.Visible = False
                lblRenewalDateRecdLink.Visible = False
                lblLongTermCareStateSpecificEffectiveDateLink.Visible = False
                lblLongTermCareStateSpecificExpirationDateLink.Visible = False

                Style.AddStyle(txtLicenseNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtRenewalDateSent, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtRenewalDateRecd, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificExpirationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtNotes, Style.StyleType.NoneditableGrayed, 400, True)

                btnEnrollerAdd.Enabled = False
                btnEnrollerRemove.Enabled = False
                btnStateAdd.Enabled = False
                btnStateRemove.Enabled = False
            End If

        Catch ex As Exception
            Throw New Exception("Error #614: BulkLicense FormatControls. " & ex.Message)
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
