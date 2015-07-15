Imports System.Data

Public Class JITAppt
    Private cCommon As Common
    Private cBulkAppointmentDataPack As BulkAppointmentDataPack

    ''''''' An appointment occurs in a state. If the qualifying criteria are met, propogate the appointment to other states ''''''' 

    Public Sub New(ByRef Page As Page)
        Try
            cCommon = New Common
            cBulkAppointmentDataPack = Page.Session("BulkAppointmentDataPack")

        Catch ex As Exception
            Throw New Exception("Error #830: JITAppt New. " & ex.Message)
        End Try
    End Sub

    Public Sub PerformExPostFactoJustInTimeAppointments(ByVal UserID As String, ByVal TriggeringState As String, ByVal LicenseEffectiveDate As String, ByVal LicenseExpirationDate As String)
        Dim i As Integer
        Dim Sql As New System.Text.StringBuilder
        Dim dtJITCarriers As DataTable
        Dim CarrierID As String
        Dim Coll As Collection
        Dim EffectiveDate As DateTime
        Dim dr As Object

        Try

            ' ___ JIT carriers
            dtJITCarriers = cCommon.GetDT("SELECT DISTINCT CarrierID FROM UserManagement..Jit_Carrier_State")

            For i = 0 To dtJITCarriers.Rows.Count - 1
                CarrierID = dtJITCarriers.Rows(i)("CarrierID")

                If IsTriggeringStateAJITStateForThisCarrier(TriggeringState, CarrierID) Then

                    Coll = IsEnrollerAppointedWithThisCarrierInResidentState(UserID, CarrierID)
                    If Coll("HasAppointment") Then

                        dr = Coll("ItemArray")
                        If CType(dr("EffectiveDate"), System.DateTime) < CType(LicenseEffectiveDate, System.DateTime) Then
                            EffectiveDate = CType(LicenseEffectiveDate, System.DateTime)
                        Else
                            EffectiveDate = CType(dr("EffectiveDate"), System.DateTime)
                        End If
                        PerformJustInTimeAppointmentsSaves(UserID, TriggeringState, CarrierID, EffectiveDate, dr("ExpirationDate"), dr("ApplicationDate"), dr("AppointmentNumber"))
                        cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.JITQualify, TriggeringState, CarrierID, 1, Nothing)
                    End If
                End If
            Next

        Catch ex As Exception
            Throw New Exception("Error #831: JITAppt PerformExPostFactoJustInTimeAppointments. " & ex.Message)
        End Try
    End Sub

    Private Function IsTriggeringStateAJITStateForThisCarrier(ByVal TriggeringState As String, ByVal CarrierID As String) As Boolean
        Dim i As Integer
        Dim dt As DataTable

        Try

            dt = cCommon.GetDT("SELECT State FROM UserManagement..JIT_Carrier_State WHERE CarrierID = '" & CarrierID & "'")
            For i = 0 To dt.Rows.Count - 1
                If dt.Rows(i)("State") = TriggeringState Then
                    Return True
                End If
            Next
            Return False

        Catch ex As Exception
            Throw New Exception("Error #832: JITAppt IsTriggeringStateAJITStateForThisCarrier. " & ex.Message)
        End Try
    End Function

    Private Function IsEnrollerAppointedWithThisCarrierInResidentState(ByVal UserID As String, ByVal CarrierID As String) As Collection
        Dim Sql As New System.Text.StringBuilder
        Dim dt As DataTable
        Dim Coll As New Collection

        Try

            Sql.Append("SELECT u.UserID, ua.ApplicationDate, ua.EffectiveDate, ua.ExpirationDate, ua.AppointmentNumber ")
            Sql.Append("FROM UserManagement..Users u ")
            Sql.Append("INNER JOIN UserManagement..UserAppointments ua ON  u.UserID = ua.UserID AND u.ResidentState = ua.State ")
            Sql.Append("WHERE u.UserID = 'jwatroba' AND ua.CarrierID = 'TRANSAMERICA' AND ISDATE(ua.EffectiveDate) = 1")
            dt = cCommon.GetDT(Sql.ToString)

            If dt.Rows.Count > 0 Then
                Coll.Add(True, "HasAppointment")
                Coll.Add(dt.Rows(0), "ItemArray")
            Else
                Coll.Add(False, "HasAppointment")
            End If
            Return Coll

        Catch ex As Exception
            Throw New Exception("Error #833: JITAppt IsEnrollerAppointedWithThisCarrierInResidentState. " & ex.Message)
        End Try
    End Function


    'For i = 0 To dtJITCarriers.Rows.Count - 1
    'Next




    Public Sub PerformJustInTimeAppointments(ByVal UserID As String, ByVal TriggeringState As String, ByVal CarrierID As String, _
            ByVal AppointmentApplicationDate As String, ByVal AppointmentNumber As String, _
            ByVal AppointmentEffectiveDate As String, ByVal AppointmentExpirationDate As String)

        Try

            Dim i As Integer
            Dim dtStateJITBlock As DataTable
            Dim JITArgs As JITArgs
            'Dim QualifyTest1Coll As Collection
            'Dim QualifyTest2Coll As Collection
            'Dim ApptApplicationDate As Object
            'Dim ApptEffectiveDate As Date
            'Dim ApptExpirationDate As Object

            ' ___ If this state is the enroller's resident state and this carrier is a JIT carrier in this state, then the trigger criteria are met
            If Not IsJITTrigger(UserID, TriggeringState, CarrierID) Then
                Exit Sub
            End If

            ' ___ Return a list of states for which this appointment triggers JIT appointments for this carrier
            dtStateJITBlock = cCommon.GetDT("SELECT State FROM UserManagement..JIT_Carrier_State WHERE CarrierID = '" & CarrierID & "'")

            For i = 0 To dtStateJITBlock.Rows.Count - 1

                JITArgs = New JITArgs

                QualifiesForJIT(JITArgs, UserID, TriggeringState, dtStateJITBlock.Rows(i)("State"), CarrierID, AppointmentEffectiveDate)
                If JITArgs.Qualifies Then

                    AdjustApptDates(JITArgs, AppointmentApplicationDate, AppointmentEffectiveDate, AppointmentExpirationDate)

                    ' ___ Delete any pending appointment for the enroller in this state for this carrier
                    If JITArgs.ApptPendingThisState Then
                        cCommon.ExecuteNonQuery("DELETE UserAppointments WHERE UserID = '" & UserID & "' AND State = '" & dtStateJITBlock.Rows(i)("State") & "' AND CarrierID = '" & CarrierID & "'")
                    End If

                    PerformJustInTimeAppointmentsSaves(UserID, dtStateJITBlock.Rows(i)("State"), CarrierID, JITArgs.ApptEffectiveDate, JITArgs.ApptExpirationDate, JITArgs.ApptApplicationDate, AppointmentNumber)
                    cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.JITQualify, dtStateJITBlock.Rows(i)("State"), CarrierID, 1, Nothing)

                Else

                    If JITArgs.Comment = "Triggering state" Or JITArgs.Comment = "Previously added" Then
                    Else
                        cBulkAppointmentDataPack.BulkAdd(BulkAppointmentDataPack.SaveTypeClass.JITQualify, dtStateJITBlock.Rows(i)("State"), CarrierID, 0, JITArgs.Comment)
                    End If

                End If
            Next

        Catch ex As Exception
            Throw New Exception("Error #834: JITAppt PerformJustInTimeAppointments. " & ex.Message)
        End Try
    End Sub

    Private Sub AdjustApptDates(ByRef JITArgs As JITArgs, ByVal AppointmentApplicationDate As String, ByVal AppointmentEffectiveDate As String, ByVal AppointmentExpirationDate As String)
        Try

            ' ___ Adjust appointment effective date to license effective date if license effective date falls after appointment effective date
            If JITArgs.LicEffectiveDate > CDate(AppointmentEffectiveDate) Then
                JITArgs.ApptEffectiveDate = JITArgs.LicEffectiveDate
            Else
                JITArgs.ApptEffectiveDate = CDate(AppointmentEffectiveDate)
            End If

            ' ___ Appointment expiration date
            If AppointmentExpirationDate.Length = 0 Then
                JITArgs.ApptExpirationDate = DBNull.Value
            Else
                JITArgs.ApptExpirationDate = CDate(AppointmentExpirationDate)
            End If

            ' ___ Application date
            JITArgs.ApptApplicationDate = DBNull.Value

            'If AppointmentApplicationDate.Length = 0 Then
            '   JITArgs.ApptApplicationDate = DBNull.Value
            'Else
            '    JITArgs.ApptApplicationDate = CDate(AppointmentApplicationDate)
            'End If

        Catch ex As Exception
            Throw New Exception("Error #835: JITAppt AdjustApptDates. " & ex.Message)
        End Try
    End Sub

    Private Sub QualifiesForJIT(ByRef JITArgs As JITArgs, ByVal UserID As String, ByVal TriggeringState As String, ByVal StateBeingTested As String, ByVal CarrierID As String, ByVal AppointmentEffectiveDate As String)
        Dim EligibleState As Boolean
        Dim EnrollerLicensedThisState As Boolean
        Dim dtLic As DataTable = Nothing
        Dim dtAppt As DataTable
        Dim PreviouslyAdded As Boolean
        Dim NoApptThisState As Boolean
        Dim ApptPendingThisState As Boolean
        Dim ApptEffectiveThisState As Boolean
        Dim Comment As String = Nothing

        Try

            ' ___ Previously added via Appointment Express?
            If cBulkAppointmentDataPack.RecordExists(StateBeingTested, CarrierID) Then
                PreviouslyAdded = True
                Comment = "Previously added"
            End If

            ' ___ Skip over the current state
            If Not PreviouslyAdded Then
                If StateBeingTested = TriggeringState Then
                    Comment = "Triggering state"
                Else
                    EligibleState = True
                End If
            End If

            ' ___ Enroller must be licensed in this state to quality for JIT
            If EligibleState Then
                dtLic = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate FROM UserManagement..UserLicenses WHERE UserID = '" & UserID & "' AND State = '" & StateBeingTested & "'")
                If dtLic.Rows.Count > 0 AndAlso cCommon.IsBVIDate(dtLic.Rows(0)("EffectiveDate")) Then
                    EnrollerLicensedThisState = True
                Else
                    Comment = "Enroller not licensed this state"
                End If
            End If

            ' ___ Enroller must not already have an appointment in this state for this carrer
            'If EligibleState And Not EnrollerLicensedThisState Then 
            If EligibleState And EnrollerLicensedThisState Then
                dtAppt = cCommon.GetDT("SELECT EffectiveDate, ExpirationDate FROM UserManagement..UserAppointments WHERE UserID = '" & UserID & "' AND CarrierID ='" & CarrierID & "' AND State = '" & StateBeingTested & "'")
                If dtAppt.Rows.Count = 0 Then
                    NoApptThisState = True
                Else
                    If IsDBNull(dtAppt.Rows(0)("EffectiveDate")) Then
                        ApptPendingThisState = True
                    Else
                        ApptEffectiveThisState = True
                        Comment = "Enroller already appointed"
                    End If
                End If
            End If

            JITArgs.Qualifies = EligibleState And EnrollerLicensedThisState And (ApptPendingThisState Or NoApptThisState)
            If JITArgs.Qualifies Then
                JITArgs.LicEffectiveDate = dtLic.Rows(0)("EffectiveDate")
                JITArgs.LicExpirationDate = dtLic.Rows(0)("ExpirationDate")
                JITArgs.ApptPendingThisState = ApptPendingThisState
            End If
            JITArgs.Comment = Comment

            'Coll.Add(Qualifies, "Qualifies")
            'If Qualifies Then
            '    Coll.Add(dtLic.Rows(0)("EffectiveDate"), "LicEffectiveDate")
            '    Coll.Add(dtLic.Rows(0)("ExpirationDate"), "LicExpirationDate")
            '    Coll.Add(ApptPendingThisState, "ApptPendingThisState")
            '    Coll.Add("JIT qualify", "Comment")
            'Else
            '    Coll.Add(Comment, "Comment")
            'End If
            'Return Coll

        Catch ex As Exception
            Throw New Exception("Error #836: JITAppt QualifiesForJIT. " & ex.Message)
        End Try
    End Sub

    Private Sub PerformJustInTimeAppointmentsSaves(ByVal UserID As String, ByVal State As String, ByVal CarrierID As String, _
            ByVal AppointmentEffectiveDate As Date, ByVal AppointmentExpirationDate As Object, ByVal AppointmentApplicationDate As Object, ByVal AppointmentNumber As String)

        Dim Sql As New System.Text.StringBuilder

        Try

            Sql.Append("INSERT INTO UserAppointments (UserID, State, ApplicationDate, EffectiveDate, ExpirationDate, CarrierID, AppointmentNumber, StatusCode, StatusCodeLastChangeDate, AddDate, ChangeDate)")
            Sql.Append(" Values ")
            Sql.Append("(" & cCommon.StrOutHandler(UserID, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(State, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("null, ")
            Sql.Append(cCommon.DateOutHandler(AppointmentEffectiveDate, True, True) & ", ")
            Sql.Append(cCommon.DateOutHandler(AppointmentExpirationDate, True, True) & ", ")
            Sql.Append(cCommon.StrOutHandler(CarrierID, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append(cCommon.StrOutHandler(AppointmentNumber, False, StringTreatEnum.SideQts_SecApost) & ", ")
            Sql.Append("'X', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "', ")
            Sql.Append("'" & cCommon.GetServerDateTime & "') ")
            cCommon.ExecuteNonQuery(Sql.ToString)

        Catch ex As Exception
            Throw New Exception("Error #837: JITAppt PerformJustInTimeAppointmentsSaves. " & ex.Message)
        End Try
    End Sub

    Public Function IsJITTrigger(ByVal UserID As String, ByVal State As String, ByVal CarrierID As String) As Boolean
        Dim dt As DataTable
        Dim ResidentState As String
        Dim JITTrigger As Boolean

        Try

            ' ___ Does this appointment occur in the appointee's resident state?
            dt = cCommon.GetDT("SELECT ResidentState FROM UserManagement..Users WHERE UserID = '" & UserID & "'")
            ResidentState = dt.Rows(0)(0)
            If ResidentState = State Then

                ' ___ Is this carrier a JIT carrier in this state?
                dt = cCommon.GetDT("SELECT Count (*) FROM UserManagement..JIT_Carrier_State WHERE CarrierID = '" & CarrierID & "' AND State = '" & State & "'")
                If dt.Rows(0)(0) > 0 Then
                    JITTrigger = True
                End If

            End If

            Return JITTrigger

        Catch ex As Exception
            Throw New Exception("Error #838: JITAppt IsJITTrigger. " & ex.Message)
        End Try
    End Function

    Public Class JITArgs
        Private cQualifies As Boolean
        Private cLicEffectiveDate As Date
        Private cLicExpirationDate As Object
        Private cApptApplicationDate As Object
        Private cApptEffectiveDate As Date
        Private cApptExpirationDate As Object
        Private cApptPendingThisState As Boolean
        Private cComment As String
        Public Property Qualifies() As Boolean
            Get
                Return cQualifies
            End Get
            Set(ByVal Value As Boolean)
                cQualifies = Value
            End Set
        End Property
        Public Property LicEffectiveDate() As Date
            Get
                Return cLicEffectiveDate
            End Get
            Set(ByVal Value As Date)
                cLicEffectiveDate = Value
            End Set
        End Property
        Public Property LicExpirationDate() As Object
            Get
                Return cLicExpirationDate
            End Get
            Set(ByVal Value As Object)
                cLicExpirationDate = Value
            End Set
        End Property
        Public Property ApptApplicationDate() As Object
            Get
                Return cApptApplicationDate
            End Get
            Set(ByVal Value As Object)
                cApptApplicationDate = Value
            End Set
        End Property
        Public Property ApptEffectiveDate() As Date
            Get
                Return cApptEffectiveDate
            End Get
            Set(ByVal Value As Date)
                cApptEffectiveDate = Value
            End Set
        End Property
        Public Property ApptExpirationDate() As Object
            Get
                Return cApptExpirationDate
            End Get
            Set(ByVal Value As Object)
                cApptExpirationDate = Value
            End Set
        End Property
        Public Property ApptPendingThisState() As Boolean
            Get
                Return cApptPendingThisState
            End Get
            Set(ByVal Value As Boolean)
                cApptPendingThisState = Value
            End Set
        End Property
        Public Property Comment() As String
            Get
                Return cComment
            End Get
            Set(ByVal Value As String)
                cComment = Value
            End Set
        End Property
    End Class
End Class
