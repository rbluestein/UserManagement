Imports System.Data
Imports System.Data.SqlClient

Partial Class UserMaintain
    Inherits System.Web.UI.Page

#Region " Declarations "
    Private cCommon As Common
    Private cEnviro As Enviro
    Private cRights As RightsClass
    Private cSess As UserWorklistSession
#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim Results As Results
        Dim RequestAction As RequestActionEnum
        Dim ResponseAction As ResponseActionEnum

        Try

            ' ___ Instantiate objects
            cCommon = New Common
            cEnviro = Session("Enviro")

            ' ___ Restore  session
            cEnviro.AuthenticateRequest(Me)

            ' ___ Right Check
            Dim RightsRqd(0) As String
            cRights = New RightsClass(Page)
            RightsRqd.SetValue(RightsClass.UserView, 0)
            cRights.HasSufficientRights(RightsRqd, True, Page)
            lblCurrentRights.Text = cCommon.GetCurRightsHidden(cRights.RightsColl)

            ' ___ Get the session object
            cSess = Session("UserWorklistSession")

            ' ___ Get RequestAction
            RequestAction = cCommon.GetRequestAction(Page)

            ' ___ Execute the RequestAction
            Results = ExecuteRequestAction(RequestAction)
            ResponseAction = Results.ResponseAction

            ' ___ Execute the ResponseAction
            If ResponseAction = ResponseActionEnum.ReturnToCallingPage Then
                cSess.PageReturnOnLoadMessage = Results.Msg
                Response.Redirect("Index.aspx?CalledBy=Child")
            Else
                DisplayPage(ResponseAction)
                If Not Results.Msg = Nothing Then
                    litMsg.Text = "<script language='javascript'>alert('" & Results.Msg & "')</script>"
                End If
                litResponseAction.Text = "<input type='hidden' name='hdResponseAction' value = '" & ResponseAction.ToString & "'>"
            End If

            ' ___ Display enviroment
            PageCaption.InnerHtml = cCommon.GetPageCaption
            'litEnviro.Text = "<input type='hidden' id='hdLoggedInUserID' name='hdLoggedInUserID' value='" & cEnviro.LoggedInUserID & "'/><input type='hidden' id='hdDBHost' name='hdDBHost'  value='" & cEnviro.DBHost & "'/>"
            litEnviro.Text = cCommon.GetlitEnviroText(cEnviro.LoggedInUserID, cEnviro.DBHost)

        Catch ex As Exception
            ' Throw New Exception("Error #150: UserMaintain Page_Load. " & ex.Message)
            Dim ErrorObj As New ErrorObj(ex, "Error #150: UserMaintain Page_Load. " & ex.Message)
        End Try
    End Sub

    Private Function ExecuteRequestAction(ByVal RequestAction As RequestActionEnum) As Results
        Dim ValidationResults As New Results
        Dim SaveResults As New Results
        Dim MyResults As New Results
        Dim TerminateAppointments As Boolean

        Try

            Select Case RequestAction
                Case RequestActionEnum.ReturnToParent
                    MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
                    MyResults.Success = True

                Case RequestActionEnum.CreateNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayBlank

                Case RequestActionEnum.SaveNew
                    ValidationResults = IsDataValid(RequestAction, TerminateAppointments)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction, TerminateAppointments)
                        If SaveResults.Success Then
                            MyResults.ResponseAction = ResponseActionEnum.ReturnToCallingPage
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
                    ValidationResults = IsDataValid(RequestAction, TerminateAppointments)
                    If ValidationResults.Success Then
                        SaveResults = PerformSave(RequestAction, TerminateAppointments)
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
                        'MyResults.ResponseAction = ResponseAction.DisplayExisting
                        MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting
                        MyResults.Msg = ValidationResults.Msg
                    End If

                Case RequestActionEnum.NoSaveNew
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputNew

                Case RequestActionEnum.NoSaveExisting
                    MyResults.ResponseAction = ResponseActionEnum.DisplayUserInputExisting

                Case RequestActionEnum.Other
                    Select Case Request.Form("hdSubAction")
                        Case "PersonalData"
                            cSess.PersonalDataSection = Not cSess.PersonalDataSection
                        Case "EmploymentData"
                            cSess.EmploymentDataSection = Not cSess.EmploymentDataSection
                        Case "Enroller"
                            cSess.EnrollerSection = Not cSess.EnrollerSection
                    End Select
                    MyResults.ResponseAction = cCommon.GetResponseActionFromRequestActionOther(Page)
            End Select

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #151: UserMaintain ExecuteRequestAction. " & ex.Message)
        End Try
    End Function

    Private Function IsDataValid(ByVal RequestAction As RequestActionEnum, ByRef TerminateAppointments As Boolean) As Results
        Dim i As Integer
        Dim MyResults As New Results
        Dim ErrColl As New Collection
        Dim dt As DataTable
        Dim IsBVI As Boolean
        Dim CurStatusCode As String

        Try

            ' ___ Trim the textbox input
            txtFirstName.Text = Trim(txtFirstName.Text)
            txtLastName.Text = Trim(txtLastName.Text)
            txtMI.Text = Trim(txtMI.Text)
            txtPrimaryContactNumber.Text = Trim(txtPrimaryContactNumber.Text)
            txtPhoneExtension.Text = Trim(txtPhoneExtension.Text)
            txtAltContactNumber.Text = Trim(txtAltContactNumber.Text)
            txtEmail.Text = Trim(txtEmail.Text)
            txtUserID.Text = Trim(txtUserID.Text)
            txtNationalProducerNumber.Text = Trim(txtNationalProducerNumber.Text)

            If chkBVI.Checked Then
                IsBVI = True
            End If

            cCommon.ValidateStringField(ErrColl, txtUserID.Text, 1, "user id not provided")
            cCommon.ValidateStringField(ErrColl, txtFirstName.Text, 1, "first name not provided")
            cCommon.ValidateStringField(ErrColl, txtLastName.Text, 1, "last name not provided")

            ' cCommon.ValidatePhoneNumber(ErrColl, txtPrimaryContactNumber.Text, "primary phone number with area code not provided")
            If txtPrimaryContactNumber.Text.Length > 0 Then
                cCommon.ValidatePhoneNumber(ErrColl, txtPrimaryContactNumber.Text, "primary contact number must be blank if valid phone number with area code not provided")
            End If
            If txtAltContactNumber.Text.Length > 0 Then
                cCommon.ValidatePhoneNumber(ErrColl, txtAltContactNumber.Text, "alt contact number must be blank if valid phone number with area code not provided")
            End If

            ' ___ Email required
            'If txtEmail.Text.Length = 0 Then
            '    cCommon.ValidateErrorOnly(ErrColl, "email address not provided")
            'Else
            '    cCommon.ValidateEmailAddress(ErrColl, txtEmail.Text, "valid email address not provided")
            'End If

            If txtEmail.Text.Length > 0 Then
                cCommon.ValidateEmailAddress(ErrColl, txtEmail.Text, "valid email address not provided")
            End If


            cCommon.ValidateDropDown(ErrColl, ddStatusCode, 1, "status not provided")
            cCommon.ValidateDropDown(ErrColl, ddResidentState, 1, "resident state not provided")


            If RequestAction = RequestActionEnum.SaveNew Then
                dt = cCommon.GetDT("Select Count (*) FROM Users WHERE lower(UserID) = '" & txtUserID.Text.ToLower & "'")
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "user id is already  in use")
                End If
            End If

            ' ___ UserID
            If RequestAction = RequestActionEnum.SaveExisting AndAlso txtUserID.Text.ToLower <> cSess.UserID.ToLower Then
                dt = cCommon.GetDT("Select Count (*) FROM Users WHERE lower(UserID) = '" & txtUserID.Text.ToLower & "'")
                If dt.Rows(0)(0) > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "user id is already  in use")
                End If
            End If

            ' ___ StatusCode
            If RequestAction = RequestActionEnum.SaveExisting Then
                dt = cCommon.GetDT("SELECT StatusCode FROM Users WHERE UserID='" & cSess.UserID & "'")
                CurStatusCode = dt.Rows(0)("StatusCode")
                'If CurStatusCode <> "INACTIVE" AndAlso ddStatusCode.SelectedValue = "INACTIVE" Then
                If CurStatusCode <> "TERMINATED" AndAlso ddStatusCode.SelectedValue = "TERMINATED" Then
                    If ddRole.SelectedValue = "ENROLLER" Or ddRole.SelectedValue = "SUPERVISOR" Then
                        If IsDate(Request.Form("hdTermDate")) Then
                            TerminateAppointments = True
                        Else
                            cCommon.ValidateErrorOnly(ErrColl, "term date not provided")
                        End If
                    End If
                End If
            End If

            If IsBVI Then
                cCommon.ValidateDropDown(ErrColl, ddLocationID, 1, "location not provided")
                cCommon.ValidateDropDown(ErrColl, ddRole, 1, "role not provided")
                cCommon.ValidateDropDown(ErrColl, ddWorkState, 1, "work state not provided")
                cCommon.ValidateDropDownSelect0(ErrColl, ddCompanyID, "company selection not permitted for BVI employee")

                cCommon.ValidatePhoneNumber(ErrColl, txtPrimaryContactNumber.Text, "primary phone number with area code not provided")
                cCommon.ValidateEmailAddress(ErrColl, txtEmail.Text, "valid email address not provided")


                If ddLongTermCareCertState.SelectedItem.Value = String.Empty Then
                    If IsDate(txtLongTermCareCertEffectiveDate.Text) Then
                        cCommon.ValidateErrorOnly(ErrColl, "long term care state cert effective date requires state cert")
                    End If
                    If IsDate(txtLongTermCareCertExpirationDate.Text) Then
                        cCommon.ValidateErrorOnly(ErrColl, "long term care state cert expiration date requires state cert")
                    End If
                Else
                    If Not IsDate(txtLongTermCareCertEffectiveDate.Text) Then
                        cCommon.ValidateErrorOnly(ErrColl, "long term care state cert requires effective date")
                    End If
                    If Not IsDate(txtLongTermCareCertExpirationDate.Text) Then
                        cCommon.ValidateErrorOnly(ErrColl, "long term care state cert requires expiration date")
                    End If
                    If IsDate(txtLongTermCareCertEffectiveDate.Text) AndAlso IsDate(txtLongTermCareCertExpirationDate.Text) Then
                        If DateTime.Compare(txtLongTermCareCertEffectiveDate.Text, txtLongTermCareCertExpirationDate.Text) > 0 Then
                            cCommon.ValidateErrorOnly(ErrColl, "ltc cert expiration date must fall after ltc cert effective date")
                        End If
                    End If

                End If

            Else
                cCommon.ValidateDropDownSelect0(ErrColl, ddLocationID, "location selection not permitted for non-BVI employee")
                cCommon.ValidateDropDownSelect0(ErrColl, ddRole, "role selection not permitted for non-BVI employee")
                cCommon.ValidateDropDownSelect0(ErrColl, ddWorkState, "workstate selection not permitted for non-BVI employee")
                If txtNationalProducerNumber.Text.Length > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "enroller national producer number not permitted for non-BVI employee")
                End If
                If ddLongTermCareCertState.SelectedItem.Value.Length > 0 Or txtLongTermCareCertEffectiveDate.Text.Length > 0 Or txtLongTermCareCertExpirationDate.Text.Length > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "enroller long term care cert info not permitted for non-BVI employee")
                End If
                cCommon.ValidateDropDown(ErrColl, ddCompanyID, 1, "company not provided")
            End If

            ' // Falcon save

            ' ___ First, refresh array in the event of a user change of Falcon data.

            ' *** Gone ***
            'HandleFalcon(ResponseActionEnum.DisplayUserInputNewOrExisting)

            ' ___ Now, validate.
            'For i = 0 To cSess.FalconData.GetUpperBound(0)
            '    If cSess.FalconData(i, 3) = String.Empty AndAlso cSess.FalconData(i, 4) <> String.Empty Then
            '        cCommon.ValidateErrorOnly(ErrColl, "Falcon " & cSess.FalconData(i, 1) & " " & cSess.FalconData(i, 2) & ": missing user id")
            '    ElseIf cSess.FalconData(i, 3) <> String.Empty AndAlso cSess.FalconData(i, 4) = String.Empty Then
            '        cCommon.ValidateErrorOnly(ErrColl, "Falcon " & cSess.FalconData(i, 1) & " " & cSess.FalconData(i, 2) & ": missing password")
            '    End If
            'Next

            If ErrColl.Count = 0 Then
                MyResults.Success = True
            Else
                MyResults.Success = False
                MyResults.Msg = cCommon.ErrCollToString(ErrColl, "Not saved. Please correct the following:")
            End If
            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #152: UserMaintain IsDataValid. " & ex.Message)
        End Try
    End Function

    Private Function PerformSave(ByVal RequestAction As RequestActionEnum, ByVal TerminateAppointments As Boolean) As Results
        Dim i As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim MyResults As New Results
        Dim IsBVI As Boolean
        Dim CurStatusCode As String = Nothing
        Dim dt As DataTable
        Dim OldUserID As String
        Dim RecID As Integer

        Try

            If chkBVI.Checked Then
                IsBVI = True
            End If

            If RequestAction = RequestActionEnum.SaveNew Then
                RecID = cCommon.GetNewRecordID("Users", "RecID")

                Sql.Append("INSERT INTO Users (RecID, FirstName, LastName, MI, PrimaryContactNumber, PhoneExtension, AltContactNumber, Email, StatusCode, ResidentState, UserID, NationalProducerNumber, LongTermCareCertState, LongTermCareCertEffectiveDate, LongTermCareCertExpirationDate, LastStatusChangeDate, Role, CompanyID, LocationID, WorkState, LastSessionID, AddDate, ChangeDate)")
                Sql.Append(" Values ")


                Sql.Append("(" & RecID.ToString & ", ")
                Sql.Append(cCommon.StrOutHandler(txtFirstName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")

                'Sql.Append("(" & cCommon.StrOutHandler(txtFirstName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtLastName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtMI.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.PhoneOutHandler(txtPrimaryContactNumber.Text, False, True) & ", ")
                Sql.Append(cCommon.PhoneOutHandler(txtPhoneExtension.Text, True, True) & ", ")
                Sql.Append(cCommon.PhoneOutHandler(txtAltContactNumber.Text, False, True) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtEmail.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddStatusCode.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(ddResidentState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtUserID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.StrOutHandler(txtNationalProducerNumber.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")

                Sql.Append(cCommon.StrOutHandler(ddLongTermCareCertState.SelectedValue, True, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtLongTermCareCertEffectiveDate.Text, True, True) & ", ")
                Sql.Append(cCommon.DateOutHandler(txtLongTermCareCertExpirationDate.Text, True, True) & ", ")

                Sql.Append("'" & cCommon.GetServerDateTime & "', ")
                If IsBVI Then
                    Sql.Append(cCommon.StrOutHandler(ddRole.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                    Sql.Append("'BVI', ")
                    Sql.Append(cCommon.StrOutHandler(ddLocationID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Else
                    Sql.Append("'CLIENT', ")
                    Sql.Append(cCommon.StrOutHandler(ddCompanyID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                    Sql.Append("'CLIENT', ")
                End If
                Sql.Append(cCommon.StrOutHandler(ddWorkState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ",")
                Sql.Append("'',")
                Sql.Append("'" & cCommon.GetServerDateTime & "',")
                Sql.Append("'" & cCommon.GetServerDateTime & "') ")

            Else

                dt = cCommon.GetDT("SELECT StatusCode FROM Users WHERE UserID = '" & cSess.UserID & "'")
                CurStatusCode = dt.Rows(0)(0)

                Sql.Append("UPDATE Users Set ")
                Sql.Append("FirstName = " & cCommon.StrOutHandler(txtFirstName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("LastName = " & cCommon.StrOutHandler(txtLastName.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("MI = " & cCommon.StrOutHandler(txtMI.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("PrimaryContactNumber = " & cCommon.PhoneOutHandler(txtPrimaryContactNumber.Text, False, True) & ", ")
                Sql.Append("PhoneExtension = " & cCommon.StrOutHandler(txtPhoneExtension.Text, True, StringTreatEnum.SideQts) & ", ")
                Sql.Append("AltContactNumber = " & cCommon.PhoneOutHandler(txtAltContactNumber.Text, False, True) & ", ")
                Sql.Append("Email = " & cCommon.StrOutHandler(txtEmail.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("StatusCode = " & cCommon.StrOutHandler(ddStatusCode.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("ResidentState = " & cCommon.StrOutHandler(ddResidentState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("UserID = " & cCommon.StrOutHandler(txtUserID.Text, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("NationalProducerNumber = " & cCommon.StrOutHandler(txtNationalProducerNumber.Text, True, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("LongTermCareCertState = " & cCommon.StrOutHandler(ddLongTermCareCertState.SelectedItem.Value, True, StringTreatEnum.SideQts_SecApost) & ", ")
                Sql.Append("LongTermCareCertEffectiveDate = " & cCommon.DateOutHandler(txtLongTermCareCertEffectiveDate.Text, True, True) & ", ")
                Sql.Append("LongTermCareCertExpirationDate = " & cCommon.DateOutHandler(txtLongTermCareCertExpirationDate.Text, True, True) & ", ")
                Sql.Append("ChangeDate = '" & cCommon.GetServerDateTime & "', ")

                If ddStatusCode.SelectedItem.Value <> CurStatusCode Then
                    Sql.Append("LastStatusChangeDate = '" & cCommon.GetServerDateTime & "', ")
                End If

                If IsBVI Then
                    Sql.Append("Role = " & cCommon.StrOutHandler(ddRole.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                    Sql.Append("CompanyID = 'BVI', ")
                    Sql.Append("LocationID = " & cCommon.StrOutHandler(ddLocationID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                Else
                    Sql.Append("Role = 'CLIENT', ")
                    Sql.Append("CompanyID = " & cCommon.StrOutHandler(ddCompanyID.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & ", ")
                    Sql.Append("LocationID = 'CLIENT', ")
                End If
                Sql.Append("WorkState = " & cCommon.StrOutHandler(ddWorkState.SelectedItem.Value, False, StringTreatEnum.SideQts_SecApost) & " ")
                Sql.Append(" WHERE UserID = '" & cSess.UserID & "'")
            End If

            Try
                cCommon.ExecuteNonQuery(Sql.ToString)
                MyResults.Success = True
                MyResults.Msg = "Update complete."
                OldUserID = cSess.UserID
                cSess.UserID = txtUserID.Text

                ' ___ Terminate appointments
                If TerminateAppointments Then
                    'cCommon.ExecuteNonQuery("UPDATE UserAppointments SET StatusCode='T', ExpirationDate = '" & Request.Form("hdTermDate") & "' WHERE UserID='" & cSess.UserID & "'")
                    'cCommon.ExecuteNonQuery("UPDATE UserAppointments SET StatusCode='T', ExpirationDate = '" & Request.Form("hdTermDate") & "' WHERE UserID='" & cSess.UserID & "' AND StatusCode <> 'T'")
                    cCommon.ExecuteNonQuery("UPDATE UserAppointments SET StatusCode='T' WHERE UserID='" & cSess.UserID & "' AND StatusCode <> 'T'")
                End If

                ' ___ Update UserID in child tables
                If RequestAction <> RequestActionEnum.SaveNew AndAlso OldUserID <> cSess.UserID Then
                    cCommon.ExecuteNonQuery("UPDATE UserLicenses SET UserID = '" & cSess.UserID & "', ChangeDate = '" & cCommon.GetServerDateTime & "' WHERE UserID = '" & OldUserID & "'")
                    cCommon.ExecuteNonQuery("UPDATE UserAppointments SET UserID = '" & cSess.UserID & "', ChangeDate = '" & cCommon.GetServerDateTime & "' WHERE UserID = '" & OldUserID & "'")
                    cCommon.ExecuteNonQuery("UPDATE UserPermissions SET UserID = '" & cSess.UserID & "', ChangeDate = '" & cCommon.GetServerDateTime & "' WHERE UserID = '" & OldUserID & "'")

                    'cCommon.ExecuteNonQuery("UPDATE UserAppointments SET UserID = '" & cSess.UserID & "' WHERE UserID = '" & OldUserID & "'")
                    'cCommon.ExecuteNonQuery("UPDATE UserPermissions SET UserID = '" & cSess.UserID & "' WHERE UserID = '" & OldUserID & "'")
                End If

                ' ___ Falcon save
                'For i = 0 To cSess.FalconData.GetUpperBound(0)
                '    If cSess.FalconData(i, 3) = String.Empty Or cSess.FalconData(i, 4) = String.Empty Then
                '        cCommon.ExecuteNonQuery("DELETE Falcon..FalconUserLookup WHERE BVIUserID = '" & cSess.UserID & "' AND CaseShortName = '" & cSess.FalconData(i, 0) & "'")
                '    Else

                '        ' ___ Test for existing record.
                '        dt = cCommon.GetDT("SELECT Count (*) FROM Falcon..FalconUserLookup WHERE BVIUserID = '" & cSess.UserID & "' AND CaseShortName = '" & cSess.FalconData(i, 0) & "'")
                '        If dt.Rows(0)(0) = 0 Then
                '            cCommon.ExecuteNonQuery("INSERT INTO Falcon..FalconUserLookup (CaseShortName, BVIUserID, FalconUserID, FalconPassword) VALUES ('" & cSess.FalconData(i, 0) & "', '" & cSess.UserID & "', '" & cSess.FalconData(i, 3) & "', '" & cSess.FalconData(i, 4) & "')")
                '        Else
                '            cCommon.ExecuteNonQuery("UPDATE Falcon..FalconUserLookup SET FalconUserID = '" & cSess.FalconData(i, 3) & "', FalconPassword = '" & cSess.FalconData(i, 4) & "' WHERE BVIUserID = '" & cSess.UserID & "' AND CaseShortName = '" & cSess.FalconData(i, 0) & "'")
                '        End If

                '    End If
                'Next

                ' ___ Notify network administrator
                If cEnviro.DBHost = "HBG-SQL" Then
                    NotifyNetAdmin(RequestAction, CurStatusCode)
                End If

            Catch ex As Exception
                MyResults.Success = False
                MyResults.Msg = "Unable to save record."
            End Try

            Return MyResults

        Catch ex As Exception
            Throw New Exception("Error #153: UserMaintain PerformSave. " & ex.Message)
        End Try
    End Function

    Private Sub NotifyNetAdmin(ByVal RequestAction As RequestActionEnum, ByVal OldStatusCode As String)
        Dim i As Integer
        Dim TotalWidth As Integer
        Dim SendTo As String
        Dim From As String
        Dim cc As String
        Dim Subject As String
        Dim sb As New System.Text.StringBuilder
        Dim Sql As String
        Dim dt As DataTable
        Dim NewStatusCode As String

        TotalWidth = 20

        ' ___ Convert status code to proper case
        OldStatusCode = cCommon.ToProper(OldStatusCode)
        NewStatusCode = cCommon.ToProper(ddStatusCode.SelectedItem.Text)

        ' ___ New or updated user
        If RequestAction = RequestActionEnum.SaveNew Then
            Subject = "New User - " & cCommon.StrOutHandler(txtUserID.Text, False, StringTreatEnum.AsIs)
        Else
            Subject = "Updated User - " & cCommon.StrOutHandler(txtUserID.Text, False, StringTreatEnum.AsIs)
        End If

        SendTo = "netadmin@benefitvision.com"
        From = "netadmin@benefitvision.com"
        cc = String.Empty

        If RequestAction = RequestActionEnum.SaveNew Then
            sb.Append(("Action type:").PadRight(TotalWidth) & "New user" & vbCrLf)
        Else
            sb.Append(("Action type:").PadRight(TotalWidth) & "Updated user" & vbCrLf)
        End If

        sb.Append(("UserID:").PadRight(TotalWidth) & cSess.UserID & vbCrLf)
        sb.Append(("LastName:").PadRight(TotalWidth) & cCommon.StrOutHandler(txtLastName.Text, False, StringTreatEnum.AsIs) & vbCrLf)
        sb.Append(("FirstName:").PadRight(TotalWidth) & cCommon.StrOutHandler(txtFirstName.Text, False, StringTreatEnum.AsIs) & vbCrLf)
        sb.Append(("LocationID:").PadRight(TotalWidth) & cCommon.StrOutHandler(ddLocationID.SelectedItem.Value, False, StringTreatEnum.AsIs) & vbCrLf)
        sb.Append(("Role:").PadRight(TotalWidth) & cCommon.StrOutHandler(ddRole.SelectedItem.Value, False, StringTreatEnum.AsIs) & vbCrLf)
        sb.Append(("StatusCode:").PadRight(TotalWidth) & NewStatusCode & vbCrLf)
        If chkBVI.Checked Then
            sb.Append(("Company:").PadRight(TotalWidth) & "BVI" & vbCrLf)
        Else
            sb.Append(("Company:").PadRight(TotalWidth) & ddCompanyID.SelectedValue & vbCrLf)
        End If
        If RequestAction <> RequestActionEnum.SaveNew Then
            sb.Append(("Old StatusCode:").PadRight(TotalWidth) & OldStatusCode & vbCrLf)
        End If
        dt = cCommon.GetDT("SELECT FirstName + ' ' + Lastname FROM Users WHERE UserID = '" & cEnviro.LoggedInUserID & "'", False)
        sb.Append(("Edited by:").PadRight(TotalWidth) & dt.Rows(0)(0))

        sb.Append(vbCrLf & vbCrLf)

        ' ___ Permissions
        sb.Append("The following permissions have been set:" & vbCrLf)

        Sql = "SELECT cc.ClientID, cc.ClientID hdClientID, cc.ClientID ClientIDStr, HasPermissionStr = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & cSess.UserID & "' and cc.ClientID = up.ClientID) = 0 then 'No' else 'Yes'  End,  HasPermission = case when (SELECT Count (*) FROM UserPermissions up WHERE UserID = '" & cSess.UserID & "' and cc.ClientID = up.ClientID) = 0 then '0' else '1'  End FROM Codes_ClientID cc ORDER BY cc.ClientID"
        dt = cCommon.GetDT(Sql, cEnviro.DBHost, "UserManagement", True)

        For i = 0 To dt.Rows.Count - 1
            ' sb.Append((dt.Rows(i)("ClientID").ToText & ":").padright(TotalWidth) & dt.Rows(i)("HasPermissionStr").ToText & vbCrLf)
            sb.Append((cCommon.Left(dt.Rows(i)("ClientID").ToText, 19) & ":").PadRight(TotalWidth) & dt.Rows(i)("HasPermissionStr").ToText & vbCrLf)
        Next

        ' cCommon.SendEmail("eodell@benefitvision.com;netadmin@benefitvision.com", cEnviro.LoggedInUserID & "@benefitvision.com", "", "User Management - New User", "Date: " & cCommon.GetServerDateTime.ToString & vbCrLf & "UserID: " & txtUserID.Text & vbCrLf & "Name: " & Trim(txtLastName.Text) & ", " & Trim(txtFirstName.Text) & vbCrLf & "Added by: " & cEnviro.LoggedInUserID)
        cCommon.SendEmail(SendTo, From, cc, Subject, sb.ToString)
    End Sub

    Private Sub DisplayPage(ByVal ResponseAction As ResponseActionEnum)
        Dim i As Integer
        Dim dt As DataTable
        Dim sb As New System.Text.StringBuilder

        Try

            ' ___ Add attributes
            chkBVI.Attributes.Add("onclick", "ToggleCompany('BVI')")
            chkOther.Attributes.Add("onclick", "ToggleCompany('Other')")
            ddStatusCode.Attributes.Add("onchange", "HandleTermDate()")
            'txtPhoneExtension.Attributes.Add("onkeyup", "fnIntegerOnly(this)")
            txtPhoneExtension.Attributes.Add("onkeyup", "fnAlphaNumericOnly(this)")

            ' ___ Heading/UserID
            Select Case ResponseAction
                Case ResponseActionEnum.DisplayBlank
                    litHeading.Text = "New User"
                Case ResponseActionEnum.DisplayUserInputNew
                    litHeading.Text = "New User"
                Case ResponseActionEnum.DisplayExisting
                    If cRights.HasThisRight(RightsClass.UserEdit) Then
                        litHeading.Text = "Edit User"
                    Else
                        litHeading.Text = "View User"
                    End If
            End Select

            If Not Page.IsPostBack Then

                ' ___ Resident/Work State/LTC Cert State
                dt = cCommon.GetDT("SELECT StateCode FROM Codes_State ORDER BY StateCode")
                ddResidentState.Items.Add(New ListItem("", ""))
                ddWorkState.Items.Add(New ListItem("", ""))
                ddLongTermCareCertState.Items.Add(New ListItem("", ""))
                For i = 0 To dt.Rows.Count - 1
                    ddResidentState.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                    ddWorkState.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                    ddLongTermCareCertState.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                Next

                ' ___ StatusCode
                dt = New DataTable
                dt = cCommon.GetDT("SELECT Status FROM Codes_Status ORDER BY Status")
                ddStatusCode.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    ddStatusCode.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                Next

                ' ___ Company
                dt = New DataTable
                dt = cCommon.GetDT("Select ClientID from Codes_ClientID Order By ClientID")
                ddCompanyID.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    If Not (dt.Rows(i)(0).ToUpper = "BVI" Or dt.Rows(i)(0).ToUpper = "BVIF") Then
                        ddCompanyID.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                    End If
                Next

                ' ___ Location
                dt = New DataTable
                dt = cCommon.GetDT("Select LocationID from Codes_LocationID Order By LocationID")
                ddLocationID.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i)(0).ToUpper <> "CLIENT" Then
                        ddLocationID.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                    End If
                Next

                ' ___ Role
                dt = New DataTable
                dt = cCommon.GetDT("Select Role from Codes_Role Order By Role")
                ddRole.Items.Add(New ListItem("", 0))
                For i = 0 To dt.Rows.Count - 1
                    If dt.Rows(i)(0).ToUpper <> "CLIENT" Then
                        ddRole.Items.Add(New ListItem(dt.Rows(i)(0), dt.Rows(i)(0)))
                    End If
                Next
            End If

            FormatControls()

            If ResponseAction = ResponseActionEnum.DisplayExisting Then

                ' ___ Get the data
                '  dt = cCommon.GetDT("Select * From Users Where UserID = '" & cSess.UserID & "'")
                dt = cCommon.GetDTSqlDataElements("SELECT * FROM Users WHERE UserID = '" & cSess.UserID & "'", "Users")
                'dtRowGuid = cCommon.GetDT("SELECT Cast(rowguid as varchar(36)) RowGuidStr FROM Users WHERE UserID = '" & cSess.UserID & "'", False)
                'cSess.ActiveRecordRowGuid = dtRowGuid.Rows(0)("RowGuidStr")

                ' ___ Populate the controls
                ddResidentState.ClearSelection()
                ddWorkState.ClearSelection()
                ddStatusCode.ClearSelection()
                ddCompanyID.ClearSelection()
                ddLocationID.ClearSelection()
                ddRole.ClearSelection()
                ddLongTermCareCertState.ClearSelection()

                ' ___ Personal data
                txtFirstName.Text = dt.Rows(0)("FirstName").ToText
                txtLastName.Text = dt.Rows(0)("LastName").ToText
                txtMI.Text = dt.Rows(0)("MI").ToText
                txtPrimaryContactNumber.Text = dt.Rows(0)("PrimaryContactNumber").ToText
                txtPhoneExtension.Text = dt.Rows(0)("PhoneExtension").ToText
                txtAltContactNumber.Text = dt.Rows(0)("AltContactNumber").ToText
                txtEmail.Text = dt.Rows(0)("Email").ToText
                ddResidentState.Items.FindByValue(dt.Rows(0)("ResidentState").ToText).Selected = True
                txtResidentState.Text = dt.Rows(0)("ResidentState").ToText

                ' ___ Employment data
                txtUserID.Text = dt.Rows(0)("UserID").ToText
                ddStatusCode.Items.FindByValue(dt.Rows(0)("StatusCode").ToText).Selected = True
                txtStatusCode.Text = dt.Rows(0)("StatusCode").ToText

                txtNationalProducerNumber.Text = dt.Rows(0)("NationalProducerNumber").ToText
                ddLongTermCareCertState.Items.FindByValue(dt.Rows(0)("LongTermCareCertState").ToText).Selected = True
                txtLongTermCareCertState.Text = dt.Rows(0)("LongTermCareCertState").ToText
                txtLongTermCareCertEffectiveDate.Text = dt.Rows(0)("LongTermCareCertEffectiveDate").ToText
                txtLongTermCareCertExpirationDate.Text = dt.Rows(0)("LongTermCareCertExpirationDate").ToText

                ' ___ Company
                If dt.Rows(0)("CompanyID").ToText.ToUpper = "BVI" Or dt.Rows(0)("CompanyID").ToText.ToUpper = "BVIF then" Then
                    chkBVI.Checked = True
                    chkOther.Checked = False
                    txtBVI.Text = "Yes"
                    txtOther.Text = "No"
                    ddLocationID.Items.FindByValue(dt.Rows(0)("LocationID").ToText).Selected = True
                    txtLocationID.Text = dt.Rows(0)("LocationID").ToText

                    ddRole.Items.FindByValue(dt.Rows(0)("Role").ToText).Selected = True
                    txtRole.Text = dt.Rows(0)("Role").ToText

                    ddWorkState.Items.FindByValue(dt.Rows(0)("WorkState").ToText).Selected = True
                    txtWorkState.Text = dt.Rows(0)("WorkState").ToText

                Else
                    chkBVI.Checked = False
                    chkOther.Checked = True
                    txtBVI.Text = "No"
                    txtOther.Text = "Yes"
                    'ddCompanyID.Items.FindByValue(dt.Rows(0)("CompanyID").ToText).Selected = True
                    cCommon.DropdownFindByValueSelect(ddCompanyID, dt.Rows(0)("CompanyID").ToText)
                    txtCompanyID.Text = dt.Rows(0)("CompanyID").ToText
                    txtStatusCode.Text = dt.Rows(0)("StatusCode").ToText
                End If

            End If

            ' *** Gone ***
            'HandleFalcon(ResponseAction)

            litHiddens.Text = "<input type='hidden' name='hdSubjUserID' value=""" & cSess.UserID & """>"

        Catch ex As Exception
            Throw New Exception("Error #154: UserMaintain DisplayPage. " & ex.Message)
        End Try
    End Sub

    'Private Sub HandleFalcon(ByVal ResponseAction As ResponseActionEnum)
    '    Dim i As Integer
    '    Dim dt As New DataTable
    '    Dim sb As New System.Text.StringBuilder
    '    Dim CaseShortName As String

    '    ' // This method has two sources. The first is IsDataValid, which requires update of the FalconData array only. 
    '    ' // It passes the argument of ResponseAction.DisplayUserInputNewOrExisting to differentiate from the second souce.
    '    ' // The second source is DisplayPage, which passes one of the other arguments. You may have noticed that
    '    ' // the arguments DisplayUserInputNew and DisplayUserInputExisting are generated by a failed save and 
    '    ' // update the array, same as DisplayUserInputNewOrExisting. This is a precaution against someone in the
    '    ' // distant future modifying the page and getting stuck on the Falcon update.

    '    Try

    '        ' ___ First, get the Falcon records.
    '        sb.Append("SELECT ccid.ClientShortName, c.CaseShortName, c.CaseDescription, FalconUserID = ISNULL(u.FalconUserID, ''), FalconPassword = ISNULL(u.FalconPassword, '') ")
    '        sb.Append("FROM Falcon..Codes_CaseShortName c ")
    '        sb.Append("INNER JOIN UserManagement..Codes_ClientID ccid ON c.ClientID = ccid.ClientID ")
    '        sb.Append("LEFT JOIN Falcon..FalconUserLookup u ON c.CaseShortName = u.CaseShortName AND u.bviuserid = '" & cSess.UserID & "' ")
    '        sb.Append("ORDER BY c.ClientID, c.CaseDescription")
    '        dt = cCommon.GetDT(sb.ToString, False)
    '        sb.Length = 0

    '        ' ___ Initialize FalconData for new or existing record
    '        Select Case ResponseAction
    '            Case ResponseActionEnum.DisplayBlank, ResponseActionEnum.DisplayExisting

    '                ' ___ Build FalconData array
    '                cSess.FalconData = Nothing
    '                If dt.Rows.Count > 0 Then
    '                    ReDim cSess.FalconData(dt.Rows.Count - 1, 4)
    '                    For i = 0 To dt.Rows.Count - 1
    '                        cSess.FalconData(i, 0) = dt.Rows(i)("CaseShortName")
    '                        cSess.FalconData(i, 1) = dt.Rows(i)("ClientShortName")
    '                        cSess.FalconData(i, 2) = dt.Rows(i)("CaseDescription")
    '                        cSess.FalconData(i, 3) = dt.Rows(i)("FalconUserID")
    '                        cSess.FalconData(i, 4) = dt.Rows(i)("FalconPassword")
    '                    Next
    '                End If

    '            Case ResponseActionEnum.DisplayUserInputNew, ResponseActionEnum.DisplayUserInputExisting, ResponseActionEnum.DisplayUserInputNewOrExisting

    '                ' // First check for page entries. These may not be present if the user has opened a previously closed enroller section. 

    '                ' ___ Update array from user input.
    '                If Request.Form("hdFalconInd") = "1" Then
    '                    For i = 0 To dt.Rows.Count - 1
    '                        cSess.FalconData(i, 3) = Request.Form(dt.Rows(i)("CaseShortName") & " _FalconUserID")
    '                        cSess.FalconData(i, 4) = Request.Form(dt.Rows(i)("CaseShortName") & " _FalconPassword")
    '                    Next
    '                End If

    '        End Select

    '        If ResponseAction <> ResponseActionEnum.DisplayUserInputNewOrExisting Then

    '            ' ___ Now, write the values to the textboxes when this method is called by DisplayPage.
    '            sb.Append("<input type = 'hidden' name = 'hdFalconInd' value = '1'><table class='PrimaryTblEmbedded' cellSpacing='0' cellPadding='0' width='100%' border='0'>")
    '            sb.Append("<tr><td>&nbsp;</td><td>Falcon UserID</td><td>Falcon Password</td></tr>")
    '            For i = 0 To cSess.FalconData.GetUpperBound(0)
    '                CaseShortName = cSess.FalconData(i, 0)
    '                sb.Append("<tr><td>" & dt.Rows(i)("ClientShortName") & " " & dt.Rows(i)("CaseDescription") & "</td>")
    '                sb.Append("<td><input type='text' onkeyup='fnAlphaNumericOnly(this)' value='" & cSess.FalconData(i, 3) & "' name='" & CaseShortName & " _FalconUserID'></td>")
    '                sb.Append("<td><input type='text' onkeyup='fnAlphaNumericOnly(this)'  value='" & cSess.FalconData(i, 4) & "' name='" & CaseShortName & " _FalconPassword'></td></tr>")
    '            Next
    '            sb.Append("</table>")
    '            litFalcon.Text = sb.ToString

    '        End If

    '    Catch ex As Exception
    '        Throw New Exception("Error #156: UserMaintain HandleFalcon. " & ex.Message)
    '    End Try
    'End Sub

    Private Sub FormatControls()
        Try

            ' ___ Show or hide sections
            plPersonalData.Visible = cSess.PersonalDataSection
            plEmploymentData.Visible = cSess.EmploymentDataSection
            plEnrollerSection.Visible = cSess.EnrollerSection

            If cRights.HasThisRight(RightsClass.UserEdit) Then
                litUpdate.Text = "<input onclick='Update()' type='button' value='Update'>"

                Style.AddStyle(txtUserID, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtFirstName, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtMI, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtLastName, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtPrimaryContactNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtPhoneExtension, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtAltContactNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtEmail, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtNationalProducerNumber, Style.StyleType.NormalEditable, 300)
                Style.AddStyle(txtLongTermCareCertEffectiveDate, Style.StyleType.NoneditableWhite, 300)
                Style.AddStyle(txtLongTermCareCertExpirationDate, Style.StyleType.NoneditableWhite, 300)

                ddResidentState.Visible = True
                txtResidentState.Visible = False
                ddStatusCode.Visible = True
                txtStatusCode.Visible = False
                ddLocationID.Visible = True
                txtLocationID.Visible = False
                ddRole.Visible = True
                txtRole.Visible = False
                ddWorkState.Visible = True
                txtWorkState.Visible = False
                ddCompanyID.Visible = True
                txtCompanyID.Visible = False
                chkBVI.Visible = True
                txtBVI.Visible = False
                chkOther.Visible = True
                txtOther.Visible = False
                ddLongTermCareCertState.Visible = True
                txtLongTermCareCertState.Visible = False
            Else

                Style.AddStyle(txtUserID, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtFirstName, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtMI, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLastName, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtPrimaryContactNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtPhoneExtension, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtAltContactNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtEmail, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtNationalProducerNumber, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareCertEffectiveDate, Style.StyleType.NoneditableGrayed, 300)
                Style.AddStyle(txtLongTermCareCertExpirationDate, Style.StyleType.NoneditableGrayed, 300)

                ddResidentState.Visible = False
                Style.AddStyle(txtResidentState, Style.StyleType.NoneditableGrayed, 300)

                ddStatusCode.Visible = False
                Style.AddStyle(txtStatusCode, Style.StyleType.NoneditableGrayed, 300)

                ddLocationID.Visible = False
                Style.AddStyle(txtLocationID, Style.StyleType.NoneditableGrayed, 300)

                ddRole.Visible = False
                Style.AddStyle(txtRole, Style.StyleType.NoneditableGrayed, 300)

                ddWorkState.Visible = False
                Style.AddStyle(txtWorkState, Style.StyleType.NoneditableGrayed, 300)

                ddCompanyID.Visible = False
                Style.AddStyle(txtCompanyID, Style.StyleType.NoneditableGrayed, 300)

                chkBVI.Visible = False
                Style.AddStyle(txtBVI, Style.StyleType.NoneditableGrayed, 300)

                chkOther.Visible = False
                Style.AddStyle(txtOther, Style.StyleType.NoneditableGrayed, 300)

                ddLongTermCareCertState.Visible = False
                Style.AddStyle(txtLongTermCareCertState, Style.StyleType.NoneditableGrayed, 300)

                lblLongTermCareCertEffectiveDateLink.Visible = False
                lblLongTermCareCertExpirationDateLink.Visible = False
            End If

        Catch ex As Exception
            Throw New Exception("Error #155: UserMaintain FormatControls. " & ex.Message)
        End Try
    End Sub
End Class
