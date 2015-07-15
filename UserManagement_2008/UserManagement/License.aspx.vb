Imports System.Data
Imports System.Data.SqlClient

Partial Class License
    Inherits System.Web.UI.Page

#Region " Declarations "
    Protected WithEvents lblFullName As System.Web.UI.WebControls.Label
    Private cEnviro As Enviro
    Private cCommon As Common
    Private cRights As RightsClass
    Private cSess As LicenseWorklistSession
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack
    Private cStateSelectionChanged As Boolean
    Private cAdminLicRight As Boolean
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

            ' ___  AdminLic specific rights
            cAdminLicRight = cRights.HasThisRight(RightsClass.AdminLicSpecific)

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
                        litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                    End If
                Else
                    litMsg.Text = "<script language='javascript'>window.showModalDialog('BulkAppointmentMessage.aspx', null, 'dialogHeight:600px;dialogWidth:570px;help:no;status:yes;scroll:yes;location:no');</script>"
                End If

                'If Not Results.Msg = Nothing Then
                '    litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                'End If



                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            litEnviro.Text = "<input type='hidden' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'><input type='hidden' name='hdDBHost'  value='" & cEnviro.DBHost & "'>"

        Catch ex As Exception
            Throw New Exception("Error #450: License Page_Load. " & ex.Message)
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

                            'MyResults.ResponseAction = ResponseAction.ReturnToCallingPage
                            'MyResults.Msg = SaveResults.Msg


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
                            MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
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
                        ' MyResults.ResponseAction = ResponseAction.DisplayExisting
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestActionEnum.NoSaveNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew

                Case RequestActionEnum.NoSaveExisting
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting

                Case RequestActionEnum.Other
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
                    If ddState.SelectedValue <> cSess.StateDropdown Then
                        cStateSelectionChanged = True
                        cSess.StateDropdown = ddState.SelectedValue
                    End If

            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #451: License ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    '' ___ Test relationship of effective and expiration dates
    'Private Sub Val_DateSequenceCheck(ByRef ErrColl As Collection, ByRef OKToProceed As Boolean)
    '    If txtEffectiveDate.Text.Length > 0 AndAlso txtExpirationDate.Text.Length > 0 Then
    '        If CType(txtEffectiveDate.Text, Date) > CType(txtExpirationDate.Text, Date) Then
    '            cCommon.ValidateErrorOnly(ErrColl, "the effective date of the license falls after the expiration date of the license")
    '            OKToProceed = False
    '        End If
    '    End If

    '    If txtEffectiveDate.Text.Length = 0 AndAlso txtExpirationDate.Text.Length > 0 Then
    '        cCommon.ValidateErrorOnly(ErrColl, "cannot have expiration date without effective date")
    '        OKToProceed = False
    '    End If
    'End Sub

    '' ___ Disallow multiple records for one state. Test for exceptions
    'Private Sub Val_LicKeyViolationCheck(ByRef RequestAction As RequestAction, ByRef ErrColl As Collection, ByRef OKToProceed As Boolean)
    '    Dim dt As DataTable

    '    If RequestAction = RequestAction.SaveNew And OKToProceed Then
    '        dt = cCommon.GetDT("SELECT Count (*) From UserLicenses WHERE UserID = '" & cSess.UserID & "' AND State = '" & ddState.SelectedItem.Value & "'")
    '        If dt.Rows(0)(0) = 1 Then
    '            cCommon.ValidateErrorOnly(ErrColl, "a license already exists for this enroller in this state")
    '            OKToProceed = False
    '        ElseIf dt.Rows(0)(0) > 1 Then
    '            cCommon.ValidateErrorOnly(ErrColl, "more than one license record exists for this enroller in this state. Please contact IT to rectify.")
    '            OKToProceed = False
    '        End If
    '    End If
    'End Sub

    '' ___ Disallow this edit if it orphans any appointments
    'Private Sub Val_LicOrphanCheck(ByRef RequestAction As RequestAction, ByRef ErrColl As Collection, ByRef OKToProceed As Boolean)
    '    Dim i As Integer
    '    Dim ApptIsCovered As Boolean
    '    Dim ApptNum As Integer
    '    Dim ApptEffDate As Date
    '    Dim ApptExpDate As Date
    '    Dim dtAppt As DataTable
    '    Dim ApptExpires As Boolean
    '    Dim ApptIsPending As Boolean
    '    Dim ApptIsActual As Boolean

    '    If RequestAction = RequestAction.SaveExisting And OKToProceed Then

    '        ' ___ Get the current license and appointment records for this user for the state.
    '        dtAppt = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate From UserAppointments WHERE UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "'")

    '        If dtAppt.Rows.Count > 0 Then
    '            For ApptNum = 0 To dtAppt.Rows.Count - 1
    '                ApptIsCovered = False
    '                ApptEffDate = Nothing
    '                ApptExpDate = Nothing
    '                ApptExpires = False
    '                ApptIsPending = False
    '                ApptIsActual = False

    '                If Not cCommon.IsBVIDate(dtAppt.Rows(ApptNum)("EffectiveDate")) Then
    '                    ApptIsPending = True
    '                Else
    '                    ApptIsActual = True
    '                    ApptEffDate = dtAppt.Rows(ApptNum)("EffectiveDate")
    '                End If

    '                If cCommon.IsBVIDate(dtAppt.Rows(ApptNum)("ExpirationDate")) Then
    '                    ApptExpires = True
    '                    ApptExpDate = dtAppt.Rows(ApptNum)("ExpirationDate")
    '                End If

    '                ApptIsCovered = Val_LicIsThisAppointmentCovered(ApptNum, ApptIsPending, ApptIsActual, ApptEffDate, ApptExpDate, ApptExpires)

    '                If Not ApptIsCovered Then
    '                    Exit For
    '                End If
    '            Next
    '            If Not ApptIsCovered Then
    '                cCommon.ValidateErrorOnly(ErrColl, "this edit will cause one or more appointments to lose license coverage")
    '            End If
    '        End If
    '    End If
    'End Sub

    'Private Function Val_LicIsThisAppointmentCovered(ByVal ApptNum As Boolean, ByVal ApptIsPending As Boolean, ByVal ApptIsActual As Boolean, ByVal ApptEffDate As Date, ByVal ApptExpDate As Date, ByVal ApptExpires As Boolean) As Boolean
    '    Dim LicEffDate As Date
    '    Dim LicExpDate As Date
    '    Dim LicExpires As Boolean
    '    Dim LicIsApplication As Boolean
    '    Dim LicIsEffective As Boolean

    '    If txtEffectiveDate.Text.Length = 0 Then
    '        LicIsApplication = True
    '    Else
    '        LicIsEffective = True
    '        LicEffDate = CDate(txtEffectiveDate.Text)
    '    End If

    '    If txtExpirationDate.Text.Length > 0 Then
    '        LicExpires = True
    '        LicExpDate = CDate(txtExpirationDate.Text)
    '    Else
    '        LicExpires = False
    '        LicExpDate = Nothing
    '    End If


    '    If ApptIsPending Then

    '        ' ___ A pending appointment is required to have a pending or actual license in this state.
    '        Return True

    '    ElseIf ApptIsActual AndAlso LicIsEffective Then

    '        ' ___ An actual appointment must be tested against an actual license and must pass one of the four tests based on effective and expiration dates.
    '        If LicExpires And ApptExpires Then
    '            If (ApptEffDate >= LicEffDate) AndAlso (ApptExpDate >= LicEffDate AndAlso ApptExpDate <= LicExpDate) Then
    '                Return True
    '            End If
    '        ElseIf LicExpires And (Not ApptExpires) Then
    '            If (ApptEffDate >= LicEffDate) AndAlso (ApptEffDate <= LicExpDate) Then
    '                Return True
    '            End If
    '        ElseIf (Not LicExpires) And ApptExpires Then
    '            If (ApptEffDate >= LicEffDate) Then
    '                Return True
    '            End If
    '        ElseIf (Not LicExpires) And (Not ApptExpires) Then
    '            If (ApptEffDate >= LicEffDate) Then
    '                Return True
    '            End If
    '        End If
    '    End If

    '    Return False
    'End Function

    Private Function IsDataValid(ByVal RequestAction As RequestActionEnum) As Results
        Dim MyResults As New Results
        Dim ErrColl As New Collection
        Dim OKToProceed As Boolean = True


        Try

            Dim LicRules As New LicRules(cCommon)

            ' ___ Trim the textbox input
            txtLicenseNumber.Text = Trim(txtLicenseNumber.Text)
            txtNotes.Text = Trim(txtNotes.Text)

            cCommon.ValidateDropDown(ErrColl, ddState, 1, "state not provided")

            If txtEffectiveDate.Text.Length > 0 And txtLicenseNumber.Text.Length = 0 Then
                cCommon.ValidateErrorOnly(ErrColl, "a license must have a license number")
            End If

            LicRules.LicAppt_LongTermCareStateSpecificDateCheck(ErrColl, cSess.StateDropdown, txtLongTermCareStateSpecificEffectiveDate, txtLongTermCareStateSpecificExpirationDate)
            LicRules.LicAppt_DateSequenceCheck(ErrColl, txtEffectiveDate, txtExpirationDate)
            LicRules.Lic_KeyViolationCheck(RequestAction, ErrColl, cSess.UserID, ddState.SelectedItem.Value)
            LicRules.Lic_AllowThisEdit(ErrColl, cSess.UserID, cSess.StateDropdown, txtEffectiveDate.Text, txtExpirationDate.Text)

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #452: License IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestActionEnum) As Results
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results
        Dim JITAppt As New JITAppt(Me)

        Try

            If RequestAction = RequestActionEnum.SaveNew Then
                Sql.Append("INSERT INTO UserLicenses (UserID, State, EffectiveDate, ExpirationDate, OKToRenewInd, LongTermCareStateSpecificEffectiveDate, LongTermCareStateSpecificExpirationDate, LicenseNumber, ApplicationDate, RenewalDateSent, RenewalDateRecd, Notes, AddDate, ChangeDate)")



                Sql.Append(" Values ")
                Sql.Append("(" & cCommon.StrOutHandler(cSess.UserID, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtEffectiveDate.Text, False, True) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")

                Sql.Append(cCommon.BitOutHandler(ddOKToRenewInd.SelectedValue, False) & ", ")

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

            ElseIf RequestAction = RequestActionEnum.SaveExisting Then

                Sql.Append("UPDATE UserLicenses Set ")
                Sql.Append("State = " & cCommon.StrOutHandler(ddState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("EffectiveDate = " & cCommon.DateOutHandler(txtEffectiveDate.Text, False, True) & ", ")
                Sql.Append("ExpirationDate = " & cCommon.DateOutHandler(txtExpirationDate.Text, True, True) & ", ")

                Sql.Append("OKToRenewInd = " & cCommon.BitOutHandler(ddOKToRenewInd.SelectedValue, False) & ", ")

                Sql.Append("LongTermCareStateSpecificEffectiveDate = " & cCommon.DateOutHandler(txtLongTermCareStateSpecificEffectiveDate.Text, True, True) & ", ")
                Sql.Append("LongTermCareStateSpecificExpirationDate = " & cCommon.DateOutHandler(txtLongTermCareStateSpecificExpirationDate.Text, True, True) & ", ")
                Sql.Append("LicenseNumber = " & cCommon.StrOutHandler(txtLicenseNumber.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")

                If IsDate(txtEffectiveDate.Text) Then
                    Sql.Append("ApplicationDate = null, ")
                Else
                    Sql.Append("ApplicationDate = " & cCommon.DateOutHandler(txtApplicationDate.Text, True, True) & ", ")
                End If

                Sql.Append("RenewalDateSent = " & cCommon.DateOutHandler(txtRenewalDateSent.Text, True, True) & ", ")
                Sql.Append("RenewalDateRecd = " & cCommon.DateOutHandler(txtRenewalDateRecd.Text, True, True) & ", ")
                Sql.Append("Notes = " & cCommon.StrOutHandler(txtNotes.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("ChangeDate = '" & cCommon.GetServerDateTime & "' ")
                Sql.Append(" WHERE UserID = '" & cSess.UserID & "' AND State ='" & cSess.State & "'")
            End If

            Try
                cCommon.ExecuteNonQuery(Sql.ToString)
                MyResults.Success = True
                MyResults.Msg = "Update complete."

                cSess.State = ddState.SelectedItem.Value
                cSess.LicenseNumber = txtLicenseNumber.Text
                cSess.EffectiveDate = txtEffectiveDate.Text
                cSess.ExpirationDate = txtExpirationDate.Text

                ' ___ Handle ExPostFacto JIT
                If RequestAction = RequestActionEnum.SaveNew AndAlso IsDate(txtEffectiveDate.Text) Then
                    JITAppt.PerformExPostFactoJustInTimeAppointments(cSess.UserID, cSess.State, txtEffectiveDate.Text, txtExpirationDate.Text)
                End If

                ' ___ Handle time-sensitive data
                cCommon.UpdateTimeSensitiveData(cSess.UserID, ddState.SelectedItem.Value)

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Msg = "Unable to save record."
            End Try

            'SqlCmd.Dispose()
            'SqlConnection1.Close()

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #453: License PerformSave. " & ex.Message)
        End Try
    End Function

    Private Sub DisplayPage(ByVal ResponseAction As ResponseActionEnum)
        Dim i As Integer
        Dim dt As DataTable
        Dim Sql As String

        Try

            ' ___ Set the attributes for controls
            txtNotes.Attributes.Add("onkeyup", "return legallength(this, 512)")
            ddState.Attributes.Add("onchange", "ClientSelectionChanged()")

            ' ___ Heading/UserID
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.DisplayUserInputNew
                    litHeading.Text = "New License"
                Case ResponseActionEnum.DisplayExisting, ResponseActionEnum.DisplayUserInputExisting
                    If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                        litHeading.Text = "Edit License"
                    Else
                        litHeading.Text = "View License"
                    End If
            End Select

            ' ___ Enroller name heading
            Sql = "SELECT LastName + ', ' +  FirstName + ' ' + MI FullName FROM Users WHERE UserID ='" & cSess.UserID & "'"
            dt = cCommon.GetDT(Sql)
            lblEnrollerName.Text = "Name:&nbsp;&nbsp;" & dt.Rows(0)("FullName")

            FormatControls(ResponseAction)

            If Not Page.IsPostBack Then

                ' ___ State
                dt = cCommon.GetDT("SELECT StateCode FROM Codes_State ORDER BY StateCode")
                ddState.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    ddState.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                Next

                ' ___ OK to renew?
                ddOKToRenewInd.Items.Add(New ListItem("Yes", "1"))
                ddOKToRenewInd.Items.Add(New ListItem("No", "0"))

            End If



            ' ___ Get the data for ResponseAction
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank
                    txtLongTermCareStateSpecificEffectiveDate.Text = String.Empty
                    txtLongTermCareStateSpecificEffectiveDate.Text = String.Empty

                    ddOKToRenewInd.SelectedValue = "1"
                    txtOKToRenewInd.Text = "Yes"

                Case ResponseActionEnum.DisplayUserInputNew
                    If cStateSelectionChanged Then
                        If cCommon.StateRequiresStateSpecificLTCCert(cSess.StateDropdown) Or cAdminLicRight Then
                            txtLongTermCareStateSpecificEffectiveDate.Text = String.Empty
                            txtLongTermCareStateSpecificExpirationDate.Text = String.Empty
                        End If
                    End If

                Case ResponseActionEnum.DisplayUserInputExisting

                Case ResponseActionEnum.DisplayExisting

                    '   EffectiveDateStr = CType(cEffectiveDate, DateTime).ToString("d")
                    'dt = Common.GetDT("SELECT * From UserLicenses Where UserID = '" & cSubjUserID & "' AND State = '" & cState & "' AND LicenseNumber='" & cLicenseNumber & "' AND EffectiveDate " & Common.DateSqlWhereNoNull(cEffectiveDate))
                    dt = cCommon.GetDT("SELECT * From UserLicenses Where UserID = '" & cSess.UserID & "' AND State = '" & cSess.State & "'")
                    ddState.ClearSelection()
                    ddState.Items.FindByValue(dt.Rows(0)("State")).Selected = True
                    txtState.Text = cCommon.StrInHandler(dt.Rows(0)("State"))
                    txtLicenseNumber.Text = cCommon.StrInHandler(dt.Rows(0)("LicenseNumber"))
                    txtApplicationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ApplicationDate"))
                    txtEffectiveDate.Text = cCommon.DateInHandler(dt.Rows(0)("EffectiveDate"))
                    txtExpirationDate.Text = cCommon.DateInHandler(dt.Rows(0)("ExpirationDate"))

                    ddOKToRenewInd.ClearSelection()
                    ddOKToRenewInd.Items.FindByValue(cCommon.BoolToBit(dt.Rows(0)("OKToRenewInd"))).Selected = True
                    txtOKToRenewInd.Text = cCommon.BoolToText(dt.Rows(0)("OKToRenewInd"))

                    txtLongTermCareStateSpecificEffectiveDate.Text = cCommon.DateInHandler(dt.Rows(0)("LongTermCareStateSpecificEffectiveDate"))
                    txtLongTermCareStateSpecificExpirationDate.Text = cCommon.DateInHandler(dt.Rows(0)("LongTermCareStateSpecificExpirationDate"))

                    txtRenewalDateSent.Text = cCommon.DateInHandler(dt.Rows(0)("RenewalDateSent"))
                    txtRenewalDateRecd.Text = cCommon.DateInHandler(dt.Rows(0)("RenewalDateRecd"))
                    txtNotes.Text = cCommon.StrInHandler(dt.Rows(0)("Notes"))
                    cSess.StateDropdown = cCommon.StrInHandler(dt.Rows(0)("State"))

            End Select

        Catch ex As Exception
            Throw New Exception("Error #454: License DisplayPage. " & ex.Message)
        End Try
    End Sub

    Private Function StateRequiresLTCCert(ByVal State As String) As Boolean
        Dim dt As DataTable
        dt = cCommon.GetDT("SELECT LTCRequiresCert FROM Codes_State WHERE upper(StateCode) = '" & State.ToUpper & "'")
        If IsDBNull(dt.Rows(0)(0)) Then
            Return False
        Else
            Return dt.Rows(0)(0)
        End If
        dt.Dispose()
    End Function


    Private Sub FormatControls(ByVal ResponseAction As ResponseActionEnum)
        Try

            If cRights.HasThisRight(RightsClass.LicenseEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"

                Style.AddStyle(txtLicenseNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableWhite, 300)

                Style.AddStyle(txtLongTermCareStateSpecificEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtLongTermCareStateSpecificExpirationDate, Style.StyleType.NoneditableWhite, 300)

                Style.AddStyle(txtRenewalDateSent, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtRenewalDateRecd, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtNotes, Style.StyleType.NormalEditable, 400, True)

                ' 6/25/08
                Select Case ResponseAction
                    Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.DisplayUserInputNew
                        ddState.Visible = True
                        txtState.Visible = False
                    Case ResponseActionEnum.DisplayExisting, ResponseActionEnum.DisplayUserInputExisting
                        ddState.Visible = False
                        Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 300)
                End Select

                ddOKToRenewInd.Visible = True
                txtOKToRenewInd.Visible = False

                lblApplicationDateLink.Visible = True
                lblEffectiveDateLink.Visible = True
                lblExpirationDateLink.Visible = True
                lblRenewalDateSentLink.Visible = True
                lblRenewalDateRecdLink.Visible = True
                lblLongTermCareStateSpecificEffectiveDateLink.Visible = True
                lblLongTermCareStateSpecificExpirationDateLink.Visible = True

                Select Case ResponseAction
                    Case ResponseActionEnum.DisplayBlank
                        plLongTermCareStateSpecific.Visible = True
                    Case Else
                        If cSess.StateDropdown = String.Empty Then
                            plLongTermCareStateSpecific.Visible = True
                        Else
                            If cCommon.StateRequiresStateSpecificLTCCert(cSess.StateDropdown) Or cAdminLicRight Then
                                plLongTermCareStateSpecific.Visible = True
                                lblLongTermCareStateSpecificEffectiveDateLink.Visible = True
                                lblLongTermCareStateSpecificExpirationDateLink.Visible = True
                            Else
                                plLongTermCareStateSpecific.Visible = False
                            End If
                        End If
                End Select

            Else

                Select Case ResponseAction
                    Case ResponseActionEnum.DisplayBlank
                        plLongTermCareStateSpecific.Visible = True
                    Case Else
                        If cSess.StateDropdown = String.Empty Then
                            plLongTermCareStateSpecific.Visible = True
                        Else
                            If cCommon.StateRequiresStateSpecificLTCCert(cSess.StateDropdown) Or cAdminLicRight Then
                                plLongTermCareStateSpecific.Visible = True
                                lblLongTermCareStateSpecificEffectiveDateLink.Visible = False
                                lblLongTermCareStateSpecificExpirationDateLink.Visible = False
                            Else
                                plLongTermCareStateSpecific.Visible = False
                            End If
                        End If
                End Select

                ddOKToRenewInd.Visible = False
                txtOKToRenewInd.Visible = True
                Style.AddStyle(txtOKToRenewInd, Style.StyleType.NoneditableGrayed, 300)

                Style.AddStyle(txtLicenseNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtApplicationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtExpirationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificExpirationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareStateSpecificExpirationDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtRenewalDateSent, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtRenewalDateRecd, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtNotes, Style.StyleType.NoneditableGrayed, 400, True)

                ddState.Visible = False
                txtState.Visible = True
                Style.AddStyle(txtState, Style.StyleType.NoneditableGrayed, 300)

                lblApplicationDateLink.Visible = False
                lblEffectiveDateLink.Visible = False
                lblExpirationDateLink.Visible = False

                lblLongTermCareStateSpecificEffectiveDateLink.Visible = False
                lblLongTermCareStateSpecificExpirationDateLink.Visible = False

                lblRenewalDateSentLink.Visible = False
                lblRenewalDateRecdLink.Visible = False
                'lblLongTermCareExpirationDateLink.Visible = False
            End If

        Catch ex As Exception
            Throw New Exception("Error #455: License FormatControls. " & ex.Message)
        End Try
    End Sub
End Class
