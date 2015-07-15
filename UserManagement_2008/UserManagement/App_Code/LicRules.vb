Imports System.Data

Public Class LicRules
    Private cCommon As Common

    Public Sub New(ByRef Common As Common)
        cCommon = Common
    End Sub


    ' ___ Is the date sequence correct?
    Public Sub LicAppt_DateSequenceCheck(ByRef ErrColl As Collection, ByVal txtEffectiveDate As TextBox, ByVal txtExpirationDate As TextBox)
        Try

            If txtEffectiveDate.Text.Length > 0 AndAlso txtExpirationDate.Text.Length > 0 Then
                If CType(txtEffectiveDate.Text, Date) > CType(txtExpirationDate.Text, Date) Then
                    cCommon.ValidateErrorOnly(ErrColl, "the effective date of the license falls after the expiration date of the license")
                End If
            End If

            If txtEffectiveDate.Text.Length = 0 AndAlso txtExpirationDate.Text.Length > 0 Then
                cCommon.ValidateErrorOnly(ErrColl, "cannot have expiration date without effective date")
            End If

        Catch ex As Exception
            Throw New Exception("Error #850: LicRules LicAppt_DateSequenceCheck. " & ex.Message)
        End Try
    End Sub

    ' ___ Is the date sequence correct?
    Public Sub LicAppt_LongTermCareStateSpecificDateCheck(ByRef ErrColl As Collection, ByVal State As String, ByVal txtEffectiveDate As TextBox, ByVal txtExpirationDate As TextBox)
        Try

            If cCommon.StateRequiresStateSpecificLTCCert(State) Then
                If IsDate(txtEffectiveDate.Text) Then
                    If IsDate(txtExpirationDate.Text) Then
                        If DateTime.Compare(txtEffectiveDate.Text, txtExpirationDate.Text) > 0 Then
                            cCommon.ValidateErrorOnly(ErrColl, "state specific expiration date must fall after state specific effective date")
                        End If
                    Else
                        cCommon.ValidateErrorOnly(ErrColl, "state specific effective date but no state specific expiration date")
                    End If
                Else
                    If IsDate(txtExpirationDate.Text) Then
                        cCommon.ValidateErrorOnly(ErrColl, "state specific expiration date but no state specific effective date")
                    End If
                End If

            Else
                If txtEffectiveDate.Text.Length > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "state specific cert not required but effective date present")
                End If
                If txtExpirationDate.Text.Length > 0 Then
                    cCommon.ValidateErrorOnly(ErrColl, "state specific cert not required but expiration date present")
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #855: LicRules LicAppt_LongTermCareStateSpecificDateCheck. " & ex.Message)
        End Try
    End Sub

    ' ___ Disallow multiple records for one state. Test for exceptions
    Public Sub Lic_KeyViolationCheck(ByRef RequestAction As RequestActionEnum, ByRef ErrColl As Collection, ByVal UserID As String, ByVal State As String)
        Dim dt As DataTable

        Try

            If RequestAction = RequestActionEnum.SaveNew Then
                dt = cCommon.GetDT("SELECT Count (*) From UserLicenses WHERE UserID = '" & UserID & "' AND State = '" & State & "'")
                If dt.Rows(0)(0) = 1 Then
                    cCommon.ValidateErrorOnly(ErrColl, "a license already exists for this enroller in this state")
                ElseIf dt.Rows(0)(0) > 1 Then
                    cCommon.ValidateErrorOnly(ErrColl, "more than one license record exists for this enroller in this state. Please contact IT to rectify.")
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #851: LicRules Lic_KeyViolationCheck. " & ex.Message)
        End Try
    End Sub

    ' ___ Disallow this edit if it orphans any appointments
    Public Sub Lic_AllowThisEdit(ByRef ErrColl As Collection, ByVal UserID As String, ByVal State As String, ByVal EffectiveDate As String, ByVal ExpirationDate As String)
        Dim i As Integer
        Dim dtAppt As DataTable
        Dim ApptIsCovered As Boolean

        Try

            Dim OrphanTestArgs As New OrphanTestArgs
            OrphanTestArgs.GetLicenseStatus(EffectiveDate, ExpirationDate)
            dtAppt = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate, StatusCode From UserAppointments WHERE UserID = '" & UserID & "' AND State = '" & State & "'")
            For i = 0 To dtAppt.Rows.Count - 1
                OrphanTestArgs.GetAppointmentStatus(dtAppt.Rows(i)("EffectiveDate"), dtAppt.Rows(i)("ExpirationDate"), dtAppt.Rows(i)("StatusCode"))
                ApptIsCovered = LicAppt_IsThisApptCoveredByThisLicense(OrphanTestArgs)
                If Not ApptIsCovered Then
                    cCommon.ValidateErrorOnly(ErrColl, "this edit will cause one or more appointments to lose license coverage")
                    Exit For
                End If
            Next

        Catch ex As Exception
            Throw New Exception("Error #852: LicRules Lic_AllowThisEdit. " & ex.Message)
        End Try
    End Sub

    Public Sub Appt_AllowThisEdit(ByRef RequestAction As RequestActionEnum, ByRef ErrColl As Collection, ByVal UserID As String, ByVal State As String, ByVal EffectiveDate As String, ByVal ExpirationDate As String, ByVal StatusCode As String)
        Dim dtLic As DataTable
        Dim ApptIsCovered As Boolean

        Try

            If RequestAction = RequestActionEnum.SaveExisting Then
                Dim OrphanTestArgs As New OrphanTestArgs
                OrphanTestArgs.GetAppointmentStatus(EffectiveDate, ExpirationDate, StatusCode)

                dtLic = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate from UserLicenses WHERE  UserID='" & UserID & "' AND  State = '" & State & "'")
                OrphanTestArgs.GetLicenseStatus(dtLic.Rows(0)("EffectiveDate"), dtLic.Rows(0)("ExpirationDate"))
                ApptIsCovered = LicAppt_IsThisApptCoveredByThisLicense(OrphanTestArgs)
                If Not ApptIsCovered Then
                    cCommon.ValidateErrorOnly(ErrColl, "unable to find a state license against which this appointment may be applied")
                End If
            End If

        Catch ex As Exception
            Throw New Exception("Error #853: LicRules Appt_AllowThisEdit. " & ex.Message)
        End Try
    End Sub

    Public Function LicAppt_IsThisApptCoveredByThisLicense(ByRef Args As OrphanTestArgs) As Boolean
        Try

            'If Args.ApptStatusCode = Args.ApptStatusCodeEnum.Terminated Then
            If Args.ApptStatusCode = OrphanTestArgs.ApptStatusCodeEnum.Terminated Then

                ' ___ Do not allow terminated appointments to invalidate license edit.
                Return True

            ElseIf Args.ApptIsPending Then

                ' ___ A pending appointment is required to have a pending or actual license in this state.
                Return True

            ElseIf Args.ApptIsActual AndAlso Args.LicIsEffective Then

                ' ___ An actual appointment must be tested against an actual license and must pass one of the four tests based on effective and expiration dates.
                If Args.LicExpires And Args.ApptExpires Then
                    If (Args.ApptEffDate >= Args.LicEffDate) AndAlso (Args.ApptExpDate >= Args.LicEffDate AndAlso Args.ApptExpDate <= Args.LicExpDate) Then
                        Return True
                    End If
                ElseIf Args.LicExpires And (Not Args.ApptExpires) Then
                    If (Args.ApptEffDate >= Args.LicEffDate) AndAlso (Args.ApptEffDate <= Args.LicExpDate) Then
                        Return True
                    End If
                ElseIf (Not Args.LicExpires) And Args.ApptExpires Then
                    If (Args.ApptEffDate >= Args.LicEffDate) Then
                        Return True
                    End If
                ElseIf (Not Args.LicExpires) And (Not Args.ApptExpires) Then
                    If (Args.ApptEffDate >= Args.LicEffDate) Then
                        Return True
                    End If
                End If

            End If

            Return False

        Catch ex As Exception
            Throw New Exception("Error #854: LicRules LicAppt_IsThisApptCoveredByThisLicense. " & ex.Message)
        End Try
    End Function

    Public Class OrphanTestArgs
        Dim cLicEffDate As Date
        Dim cLicExpDate As Date
        Dim cLicExpires As Boolean
        Dim cLicIsApplication As Boolean
        Dim cLicIsEffective As Boolean

        Dim cApptEffDate As Date
        Dim cApptExpDate As Date
        Dim cApptExpires As Boolean
        Dim cApptIsPending As Boolean
        Dim cApptIsActual As Boolean
        Dim cApptStatusCode As ApptStatusCodeEnum

        Public Enum ApptStatusCodeEnum
            Pending = 1
            Effective = 2
            Terminated = 3
        End Enum


        Public ReadOnly Property LicEffDate() As Date
            Get
                Return cLicEffDate
            End Get
        End Property
        Public ReadOnly Property LicExpDate() As Date
            Get
                Return cLicExpDate
            End Get
        End Property
        Public ReadOnly Property LicExpires() As Boolean
            Get
                Return cLicExpires
            End Get
        End Property
        Public ReadOnly Property LicIsAppLication() As Boolean
            Get
                Return cLicIsApplication
            End Get
        End Property
        Public ReadOnly Property LicIsEffective() As Boolean
            Get
                Return cLicIsEffective
            End Get
        End Property

        Public Property ApptEffDate() As Date
            Get
                Return cApptEffDate
            End Get
            Set(ByVal Value As Date)
                cApptEffDate = Value
            End Set
        End Property
        Public ReadOnly Property ApptExpDate() As Date
            Get
                Return cApptExpDate
            End Get
        End Property
        Public ReadOnly Property ApptExpires() As Boolean
            Get
                Return cApptExpires
            End Get
        End Property

        Public ReadOnly Property ApptIsPending() As Boolean
            Get
                Return cApptIsPending
            End Get
        End Property
        Public ReadOnly Property ApptIsActual() As Boolean
            Get
                Return cApptIsActual
            End Get
        End Property
        Public ReadOnly Property ApptStatusCode() As ApptStatusCodeEnum
            Get
                Return cApptStatusCode
            End Get
        End Property





        Public Sub GetLicenseStatus(ByVal LicEffDate As Object, ByVal LicExpDate As Object)
            If IsBVIDate(LicEffDate) Then
                cLicIsEffective = True
                cLicEffDate = CDate(LicEffDate)
            Else
                cLicIsApplication = True
            End If
            If IsBVIDate(LicExpDate) Then
                cLicExpires = True
                cLicExpDate = CDate(LicExpDate)
            Else
                cLicIsApplication = True
            End If
        End Sub

        'Public Sub GetAppointmentStatus(ByVal ApptEffDate As Object, ByVal ApptExpDate As Object)
        '    If Not IsBVIDate(ApptEffDate) Then
        '        cApptIsPending = True
        '    Else
        '        cApptIsActual = True
        '        cApptEffDate = CDate(ApptEffDate)
        '    End If

        '    If IsBVIDate(ApptExpDate) Then
        '        cApptExpires = True
        '        cApptExpDate = CDate(ApptExpDate)
        '    End If
        'End Sub

        Public Sub GetAppointmentStatus(ByVal ApptEffDate As Object, ByVal ApptExpDate As Object, ByVal StatusCode As String)
            If Not IsBVIDate(ApptEffDate) Then
                cApptIsPending = True
            Else
                cApptIsActual = True
                cApptEffDate = CDate(ApptEffDate)
            End If

            If IsBVIDate(ApptExpDate) Then
                cApptExpires = True
                cApptExpDate = CDate(ApptExpDate)
            End If

            Select Case StatusCode
                Case "P"
                    cApptStatusCode = ApptStatusCodeEnum.Pending
                Case "X"
                    cApptStatusCode = ApptStatusCodeEnum.Effective
                Case "T"
                    cApptStatusCode = ApptStatusCodeEnum.Terminated
            End Select
        End Sub

        Private Function IsBVIDate(ByVal Input As Object) As Boolean
            If IsDBNull(Input) Then
                Return False
            ElseIf Input = Nothing Then
                Return False
            ElseIf Input = "01/01/1950" Then
                Return False
            ElseIf Input.ToString = String.Empty Then
                Return False
            Else
                Return True
            End If
        End Function
    End Class


    Public Sub Appt_TestOneCarrierPerState(ByVal RequestAction As RequestActionEnum, ByRef ErrColl As Collection, ByVal Sess As LicenseWorklistSession, ByVal CarrierID As String)
        Dim dt As DataTable

        If RequestAction = RequestActionEnum.SaveNew Then
            dt = cCommon.GetDT("SELECT Count (*) From UserAppointments WHERE UserID = '" & Sess.UserID & "' AND CarrierID = '" & CarrierID & "'  AND State = '" & Sess.State & "'")
            If dt.Rows(0)(0) > 0 Then
                cCommon.ValidateErrorOnly(ErrColl, "this enroller already has an appointment record with this carrier in this state.")
            End If
            dt = Nothing
        ElseIf RequestAction = RequestActionEnum.SaveExisting Then
            ' No action
        End If
    End Sub
End Class
