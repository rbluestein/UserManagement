Imports System.Data
Imports System.Data.SqlClient

Public Class PageSession
    Private cSortReference As String = String.Empty
    Private cFilterOnOffState As String = String.Empty
    Private cPageInitiallyLoaded As Boolean
    Private cPageReturnOnLoadMessasge As String = String.Empty

    Public Property SortReference() As String
        Get
            Return cSortReference
        End Get
        Set(ByVal Value As String)
            cSortReference = Value
        End Set
    End Property
    Public Property PageReturnOnLoadMessage() As String
        Get
            Return cPageReturnOnLoadMessasge
        End Get
        Set(ByVal Value As String)
            cPageReturnOnLoadMessasge = Value
        End Set
    End Property
    Public Property FilterOnOffState() As String
        Get
            Return cFilterOnOffState
        End Get
        Set(ByVal Value As String)
            cFilterOnOffState = Value
        End Set
    End Property
    Public Property PageInitiallyLoaded() As Boolean
        Get
            Return cPageInitiallyLoaded
        End Get
        Set(ByVal Value As Boolean)
            cPageInitiallyLoaded = Value
        End Set
    End Property
End Class

Public Class CompanyWorklistSession
    Inherits PageSession

    Private cCompanyID As String = String.Empty
    Private cCompanyIDSelectedFilterValue As String = String.Empty
    Private cPrimaryContactNameSelectedFilterValue As String = String.Empty

    Public Property CompanyID() As String
        Get
            Return cCompanyID
        End Get
        Set(ByVal Value As String)
            cCompanyID = Value
        End Set
    End Property
    Public Property CompanyIDSelectedFilterValue() As String
        Get
            Return cCompanyIDSelectedFilterValue
        End Get
        Set(ByVal Value As String)
            cCompanyIDSelectedFilterValue = Value
        End Set
    End Property
    Public Property PrimaryContactNameSelectedFilterValue() As String
        Get
            Return cPrimaryContactNameSelectedFilterValue
        End Get
        Set(ByVal Value As String)
            cPrimaryContactNameSelectedFilterValue = Value
        End Set
    End Property
End Class

Public Class ClientWorklistSession
    Inherits PageSession

    Private cClientID As String = String.Empty
    Private cClientIDSelectedFilterValue As String = String.Empty

    Public Property ClientID() As String
        Get
            Return cClientID
        End Get
        Set(ByVal Value As String)
            cClientID = Value
        End Set
    End Property
    Public Property ClientIDSelectedFilterValue() As String
        Get
            Return cClientIDSelectedFilterValue
        End Get
        Set(ByVal Value As String)
            cClientIDSelectedFilterValue = Value
        End Set
    End Property
End Class

Public Class CarrierWorklistSession
    Inherits PageSession

    Private cCarrierID As String = String.Empty
    Private cCarrierIDSelectedFilterValue As String = String.Empty

    Public Property CarrierID() As String
        Get
            Return cCarrierID
        End Get
        Set(ByVal Value As String)
            cCarrierID = Value
        End Set
    End Property
    Public Property CarrierIDSelectedFilterValue() As String
        Get
            Return cCarrierIDSelectedFilterValue
        End Get
        Set(ByVal Value As String)
            cCarrierIDSelectedFilterValue = Value
        End Set
    End Property
End Class

Public Class UserWorklistSession
    Inherits PageSession

    Private cUserID As String = String.Empty
    Private cUserIDFilter As String = String.Empty
    Private cFullNameFilter As String = String.Empty
    Private cStatusCodeFilter As String = String.Empty
    Private cRoleFilter As String = String.Empty
    Private cCompanyIDFilter As String = String.Empty
    Private cLocationIDFilter As String = String.Empty
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cSql As String
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cSessionID As String
    Private cActiveRecordRowGuid As String
    Private cEnrollerSection As Boolean = True
    Private cPersonalDataSection As Boolean = True
    Private cEmploymentDataSection As Boolean = True
    Private cFalconData As Object

    Public Sub New()
        MyBase.new()
        cSessionID = Guid.NewGuid.ToString
    End Sub

    Public ReadOnly Property SessionID() As String
        Get
            Return cSessionID
        End Get
    End Property
    Public Property UserID() As String
        Get
            Return cUserID
        End Get
        Set(ByVal Value As String)
            cUserID = Value
        End Set
    End Property
    Public Property UserIDFilter() As String
        Get
            Return cUserIDFilter
        End Get
        Set(ByVal Value As String)
            cUserIDFilter = Value
        End Set
    End Property
    Public Property FullNameFilter() As String
        Get
            Return cFullNameFilter
        End Get
        Set(ByVal Value As String)
            cFullNameFilter = Value
        End Set
    End Property
    Public Property StatusCodeFilter() As String
        Get
            Return cStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cStatusCodeFilter = Value
        End Set
    End Property
    Public Property RoleFilter() As String
        Get
            Return cRoleFilter
        End Get
        Set(ByVal Value As String)
            cRoleFilter = Value
        End Set
    End Property
    Public Property CompanyIDFilter() As String
        Get
            Return cCompanyIDFilter
        End Get
        Set(ByVal Value As String)
            cCompanyIDFilter = Value
        End Set
    End Property
    Public Property LocationIDFilter() As String
        Get
            Return cLocationIDFilter
        End Get
        Set(ByVal Value As String)
            cLocationIDFilter = Value
        End Set
    End Property
    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property
    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property
    Public Property ActiveRecordRowGuid() As String
        Get
            Return cActiveRecordRowGuid
        End Get
        Set(ByVal Value As String)
            cActiveRecordRowGuid = Value
        End Set
    End Property
    Public Property EnrollerSection() As Boolean
        Get
            Return cEnrollerSection
        End Get
        Set(ByVal Value As Boolean)
            cEnrollerSection = Value
        End Set
    End Property
    Public Property EmploymentDataSection() As Boolean
        Get
            Return cEmploymentDataSection
        End Get
        Set(ByVal Value As Boolean)
            cEmploymentDataSection = Value
        End Set
    End Property
    Public Property PersonalDataSection() As Boolean
        Get
            Return cPersonalDataSection
        End Get
        Set(ByVal Value As Boolean)
            cPersonalDataSection = Value
        End Set
    End Property
    Public Property FalconData() As Object
        Get
            Return cFalconData
        End Get
        Set(ByVal Value As Object)
            cFalconData = Value
        End Set
    End Property
End Class

Public Class LicenseWorklistSession
    Inherits PageSession

    ' ___ Active fields
    Private cUserID As String = String.Empty
    Private cLicenseNumber As String = String.Empty
    Private cEffectiveDate As String = String.Empty
    Private cExpirationDate As String = String.Empty
    Private cState As String = String.Empty
    Private cCarrierID As String = String.Empty
    Private cAppointmentNumber As String = String.Empty

    ' ___ Subtable selection
    Private cSubTableInd As String = String.Empty
    Private cSubTableLicenseNumber As String = String.Empty
    Private cSubTableEffectiveDate As String = String.Empty
    Private cSubTableTableCarrierID As String = String.Empty
    Private cSubTableState As String = String.Empty

    ' ___ Filter selections
    Private cStateFilter As String = String.Empty
    Private cCatgyFilter As String = String.Empty

    ' ___ License state dropdown
    Private cStateDropdown As String = String.Empty

    Public Property StateDropdown() As String
        Get
            Return cStateDropdown
        End Get
        Set(ByVal Value As String)
            cStateDropdown = Value
        End Set
    End Property
    Public Property AppointmentNumber() As String
        Get
            Return cAppointmentNumber
        End Get
        Set(ByVal Value As String)
            cAppointmentNumber = Value
        End Set
    End Property

    Public Property UserID() As String
        Get
            Return cUserID
        End Get
        Set(ByVal Value As String)
            cUserID = Value
        End Set
    End Property
    Public Property LicenseNumber() As String
        Get
            Return cLicenseNumber
        End Get
        Set(ByVal Value As String)
            cLicenseNumber = Value
        End Set
    End Property
    Public Property EffectiveDate() As String
        Get
            Return cEffectiveDate
        End Get
        Set(ByVal Value As String)
            cEffectiveDate = Value
        End Set
    End Property
    Public Property ExpirationDate() As String
        Get
            Return cExpirationDate
        End Get
        Set(ByVal Value As String)
            cExpirationDate = Value
        End Set
    End Property
    Public Property State() As String
        Get
            Return cState
        End Get
        Set(ByVal Value As String)
            cState = Value
        End Set
    End Property
    Public Property CarrierID() As String
        Get
            Return cCarrierID
        End Get
        Set(ByVal Value As String)
            cCarrierID = Value
        End Set
    End Property

    Public Property SubTableInd() As String
        Get
            Return cSubTableInd
        End Get
        Set(ByVal Value As String)
            cSubTableInd = Value
        End Set
    End Property
    Public Property SubTableLicenseNumber() As String
        Get
            Return cSubTableLicenseNumber
        End Get
        Set(ByVal Value As String)
            cSubTableLicenseNumber = Value
        End Set
    End Property
    Public Property SubTableEffectiveDate() As String
        Get
            Return cSubTableEffectiveDate
        End Get
        Set(ByVal Value As String)
            cSubTableEffectiveDate = Value
        End Set
    End Property
    Public Property SubTableTableCarrierID() As String
        Get
            Return cSubTableTableCarrierID
        End Get
        Set(ByVal Value As String)
            cSubTableTableCarrierID = Value
        End Set
    End Property
    Public Property SubTableState() As String
        Get
            Return cSubTableState
        End Get
        Set(ByVal Value As String)
            cSubTableState = Value
        End Set
    End Property
    Public Property StateFilter() As String
        Get
            Return cStateFilter
        End Get
        Set(ByVal Value As String)
            cStateFilter = Value
        End Set
    End Property
    Public Property CatgyFilter() As String
        Get
            Return cCatgyFilter
        End Get
        Set(ByVal Value As String)
            cCatgyFilter = Value
        End Set
    End Property
End Class

Public Class BulkLicSession
    Inherits PageSession


    Private cInitialized As Boolean
    Private cLetter As String
    Private cUserId As String = String.Empty
    Private cStateTarget As String = String.Empty
    Private cStateFullTarget As String = String.Empty

    Private cLicenseNumber As String = String.Empty
    Private cLongTermCareStateSpecificEffectiveDate As Date
    Private cLongTermCareStateSpecificExpirationDate As Date


    Private cNotes As String = String.Empty
    Private cApplicationDate As Date
    Private cEffectiveDate As Date
    Private cExpirationDate As Date
    Private cRenewalDateSent As Date
    Private cRenewalDateRecd As Date
    'Private cLongTermCareExpirationDate As Date
    Private cEnrollerTarget As String = String.Empty

    Public Property LicenseNumber() As String
        Get
            Return cLicenseNumber
        End Get
        Set(ByVal Value As String)
            cLicenseNumber = Value
        End Set
    End Property
    Public Property LongTermCareStateSpecificEffectiveDate() As Date
        Get
            Return cLongTermCareStateSpecificEffectiveDate
        End Get
        Set(ByVal Value As Date)
            cLongTermCareStateSpecificEffectiveDate = Value
        End Set
    End Property
    Public Property LongTermCareStateSpecificExpirationDate() As Date
        Get
            Return cLongTermCareStateSpecificExpirationDate
        End Get
        Set(ByVal Value As Date)
            cLongTermCareStateSpecificExpirationDate = Value
        End Set
    End Property
    Public Property Notes() As String
        Get
            Return cNotes
        End Get
        Set(ByVal Value As String)
            cNotes = Value
        End Set
    End Property
    Public Property ApplicationDate() As Date
        Get
            Return cApplicationDate
        End Get
        Set(ByVal Value As Date)
            cApplicationDate = Value
        End Set
    End Property
    Public Property EffectiveDate() As Date
        Get
            Return cEffectiveDate
        End Get
        Set(ByVal Value As Date)
            cEffectiveDate = Value
        End Set
    End Property
    Public Property RenewalDateSent() As Date
        Get
            Return cRenewalDateSent
        End Get
        Set(ByVal Value As Date)
            cRenewalDateSent = Value
        End Set
    End Property
    Public Property RenewalDateRecd() As Date
        Get
            Return cRenewalDateRecd
        End Get
        Set(ByVal Value As Date)
            cRenewalDateRecd = Value
        End Set
    End Property
    Public Property ExpirationDate() As Date
        Get
            Return cExpirationDate
        End Get
        Set(ByVal Value As Date)
            cExpirationDate = Value
        End Set
    End Property
    Public Property Initialized() As Boolean
        Get
            Return cInitialized
        End Get
        Set(ByVal Value As Boolean)
            cInitialized = Value
        End Set
    End Property
    Public Property Letter() As String
        Get
            Return cLetter
        End Get
        Set(ByVal Value As String)
            cLetter = Value
        End Set
    End Property
    Public Property UserId() As String
        Get
            Return cUserId
        End Get
        Set(ByVal Value As String)
            cUserId = Value
        End Set
    End Property
    Public Property StateTarget() As String
        Get
            Return cStateTarget
        End Get
        Set(ByVal Value As String)
            cStateTarget = Value
        End Set
    End Property
    Public Property StateFullTarget() As String
        Get
            Return cStateFullTarget
        End Get
        Set(ByVal Value As String)
            cStateFullTarget = Value
        End Set
    End Property
    Public Property EnrollerTarget() As String
        Get
            Return cEnrollerTarget
        End Get
        Set(ByVal Value As String)
            cEnrollerTarget = Value
        End Set
    End Property

End Class

Public Class BulkApptSession
    Inherits PageSession

    Private cInitialized As Boolean
    Private cLetter As String
    Private cUserId As String = String.Empty
    Private cEnrollerTarget As String = String.Empty

    Private cCarrierTarget As String = String.Empty
    Private cCarrierTargetList As New SortedList

    Private cStateTarget As String = String.Empty
    Private cStateTargetList As New SortedList

    Private cAppointmentNumber As String = String.Empty
    Private cApplicationDate As Date
    Private cEffectiveDate As Date
    Private cExpirationDate As Date

    Private cStateSourceSelectedState As String = String.Empty
    Private cStatusCode As String = String.Empty


    Public Property StateSourceSelectedState() As String
        Get
            Return cStateSourceSelectedState
        End Get
        Set(ByVal Value As String)
            cStateSourceSelectedState = Value
        End Set
    End Property

    Public Property AppointmentNumber() As String
        Get
            Return cAppointmentNumber
        End Get
        Set(ByVal Value As String)
            cAppointmentNumber = Value
        End Set
    End Property

    Public Property ApplicationDate() As Date
        Get
            Return cApplicationDate
        End Get
        Set(ByVal Value As Date)
            cApplicationDate = Value
        End Set
    End Property
    Public Property EffectiveDate() As Date
        Get
            Return cEffectiveDate
        End Get
        Set(ByVal Value As Date)
            cEffectiveDate = Value
        End Set
    End Property
    Public Property ExpirationDate() As Date
        Get
            Return cExpirationDate
        End Get
        Set(ByVal Value As Date)
            cExpirationDate = Value
        End Set
    End Property
    Public Property Initialized() As Boolean
        Get
            Return cInitialized
        End Get
        Set(ByVal Value As Boolean)
            cInitialized = Value
        End Set
    End Property
    Public Property Letter() As String
        Get
            Return cLetter
        End Get
        Set(ByVal Value As String)
            cLetter = Value
        End Set
    End Property
    Public Property UserId() As String
        Get
            Return cUserId
        End Get
        Set(ByVal Value As String)
            cUserId = Value
        End Set
    End Property
    Public Property CarrierTarget() As String
        Get
            Return cCarrierTarget
        End Get
        Set(ByVal Value As String)
            cCarrierTarget = Value
        End Set
    End Property
    Public ReadOnly Property CarrierTargetList() As SortedList
        Get
            Return cCarrierTargetList
        End Get
    End Property
    Public Property StateTarget() As String
        Get
            Return cStateTarget
        End Get
        Set(ByVal Value As String)
            cStateTarget = Value
        End Set
    End Property
    Public ReadOnly Property StateTargetList() As SortedList
        Get
            Return cStateTargetList
        End Get
    End Property
    Public Property EnrollerTarget() As String
        Get
            Return cEnrollerTarget
        End Get
        Set(ByVal Value As String)
            cEnrollerTarget = Value
        End Set
    End Property
    Public Property StatusCode() As String
        Get
            Return cStatusCode
        End Get
        Set(ByVal Value As String)
            cStatusCode = Value
        End Set
    End Property
End Class

Public Class QueryEnrollerSession
    Inherits PageSession

    Private cUserID As String = String.Empty
    Private cUserIDFilter As String = String.Empty
    Private cFullNameFilter As String = String.Empty
    Private cEmpStatusCodeFilter As String = String.Empty
    Private cLocationIDFilter As String = String.Empty
    Private cLicStatus As String = String.Empty
    Private cCarStatus As String = String.Empty
    Private cStateFilter As String = String.Empty
    Private cFilterCarrier As String = String.Empty
    Private cSessionID As String
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cSql As String = String.Empty


    Public Sub New()
        MyBase.new()
        cSessionID = Guid.NewGuid.ToString
    End Sub

    Public ReadOnly Property SessionID() As String
        Get
            Return cSessionID
        End Get
    End Property
    Public Property UserID() As String
        Get
            Return cUserID
        End Get
        Set(ByVal Value As String)
            cUserID = Value
        End Set
    End Property
    Public Property UserIDFilter() As String
        Get
            Return cUserIDFilter
        End Get
        Set(ByVal Value As String)
            cUserIDFilter = Value
        End Set
    End Property
    Public Property FullNameFilter() As String
        Get
            Return cFullNameFilter
        End Get
        Set(ByVal Value As String)
            cFullNameFilter = Value
        End Set
    End Property
    Public Property EmpStatusCodeFilter() As String
        Get
            Return cEmpStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cEmpStatusCodeFilter = Value
        End Set
    End Property
    Public Property LocationIDFilter() As String
        Get
            Return cLocationIDFilter
        End Get
        Set(ByVal Value As String)
            cLocationIDFilter = Value
        End Set
    End Property
    Public Property LicStatus() As String
        Get
            Return cLicStatus
        End Get
        Set(ByVal Value As String)
            cLicStatus = Value
        End Set
    End Property
    Public Property CarStatus() As String
        Get
            Return cCarStatus
        End Get
        Set(ByVal Value As String)
            cCarStatus = Value
        End Set
    End Property
    Public Property FilterCarrier() As String
        Get
            Return cFilterCarrier
        End Get
        Set(ByVal Value As String)
            cFilterCarrier = Value
        End Set
    End Property
    Public Property StateFilter() As String
        Get
            Return cStateFilter
        End Get
        Set(ByVal Value As String)
            cStateFilter = Value
        End Set
    End Property
    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property
End Class

Public Class QueryCarrierSession
    Inherits PageSession

    Private cUserID As String = String.Empty
    Private cUserIDFilter As String = String.Empty
    Private cFullNameFilter As String = String.Empty
    Private cEmpStatusCodeFilter As String = String.Empty
    Private cLocationIDFilter As String = String.Empty
    Private cLicStatusFilter As String = String.Empty
    Private cCarStatusCodeFilter As String = String.Empty
    Private cStateFilter As String = String.Empty
    Private cCarrierIDFilter As String = String.Empty
    Private cSessionID As String
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cSql As String = String.Empty

    Public ReadOnly Property SessionID() As String
        Get
            Return cSessionID
        End Get
    End Property
    Public Property UserID() As String
        Get
            Return cUserID
        End Get
        Set(ByVal Value As String)
            cUserID = Value
        End Set
    End Property
    Public Property UserIDFilter() As String
        Get
            Return cUserIDFilter
        End Get
        Set(ByVal Value As String)
            cUserIDFilter = Value
        End Set
    End Property
    Public Property FullNameFilter() As String
        Get
            Return cFullNameFilter
        End Get
        Set(ByVal Value As String)
            cFullNameFilter = Value
        End Set
    End Property
    Public Property EmpStatusCodeFilter() As String
        Get
            Return cEmpStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cEmpStatusCodeFilter = Value
        End Set
    End Property
    Public Property LocationIDFilter() As String
        Get
            Return cLocationIDFilter
        End Get
        Set(ByVal Value As String)
            cLocationIDFilter = Value
        End Set
    End Property
    Public Property LicStatusFilter() As String
        Get
            Return cLicStatusFilter
        End Get
        Set(ByVal Value As String)
            cLicStatusFilter = Value
        End Set
    End Property
    Public Property CarStatusCodeFilter() As String
        Get
            Return cCarStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cCarStatusCodeFilter = Value
        End Set
    End Property
    Public Property StateFilter() As String
        Get
            Return cStateFilter
        End Get
        Set(ByVal Value As String)
            cStateFilter = Value
        End Set
    End Property
    Public Property CarrierIDFilter() As String
        Get
            Return cCarrierIDFilter
        End Get
        Set(ByVal Value As String)
            cCarrierIDFilter = Value
        End Set
    End Property
    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property

    Public Sub New()
        MyBase.new()
        cSessionID = Guid.NewGuid.ToString
    End Sub
End Class

Public Class QueryLicenseSession
    Inherits PageSession

    Private cUserID As String = String.Empty
    Private cUserIDFilter As String = String.Empty
    Private cFullNameFilter As String = String.Empty
    Private cEmpStatusCodeFilter As String = String.Empty
    Private cLocationIDFilter As String = String.Empty
    Private cStateFilter As String = String.Empty
    Private cCatgyFilter As String = String.Empty
    Private cExpirationDateFromFilter As String = String.Empty
    Private cExpirationDateToFilter As String = String.Empty
    Private cOKToRenewIndFilter As String = String.Empty
    Private cExcessiveRecordsWarningInEffect As Boolean
    Private cSql As String = String.Empty
    Private cDateRange As String
    Private cInitialReportDataSuppressInEffect As Boolean
    Private cSessionID As String

    Public Sub New()
        MyBase.new()
        cSessionID = Guid.NewGuid.ToString
    End Sub

    Public Property UserID() As String
        Get
            Return cUserID
        End Get
        Set(ByVal Value As String)
            cUserID = Value
        End Set
    End Property
    Public Property UserIDFilter() As String
        Get
            Return cUserIDFilter
        End Get
        Set(ByVal Value As String)
            cUserIDFilter = Value
        End Set
    End Property
    Public Property FullNameFilter() As String
        Get
            Return cFullNameFilter
        End Get
        Set(ByVal Value As String)
            cFullNameFilter = Value
        End Set
    End Property
    Public Property EmpStatusCodeFilter() As String
        Get
            Return cEmpStatusCodeFilter
        End Get
        Set(ByVal Value As String)
            cEmpStatusCodeFilter = Value
        End Set
    End Property
    Public Property LocationIDFilter() As String
        Get
            Return cLocationIDFilter
        End Get
        Set(ByVal Value As String)
            cLocationIDFilter = Value
        End Set
    End Property
    Public Property StateFilter() As String
        Get
            Return cStateFilter
        End Get
        Set(ByVal Value As String)
            cStateFilter = Value
        End Set
    End Property
    Public Property CatgyFilter() As String
        Get
            Return cCatgyFilter
        End Get
        Set(ByVal Value As String)
            cCatgyFilter = Value
        End Set
    End Property
    Public Property ExpirationDateFromFilter() As String
        Get
            Return cExpirationDateFromFilter
        End Get
        Set(ByVal Value As String)
            cExpirationDateFromFilter = Value
        End Set
    End Property
    Public Property ExpirationDateToFilter() As String
        Get
            Return cExpirationDateToFilter
        End Get
        Set(ByVal Value As String)
            cExpirationDateToFilter = Value
        End Set
    End Property
    Public Property OKToRenewIndFilter() As String
        Get
            Return cOKToRenewIndFilter
        End Get
        Set(ByVal Value As String)
            cOKToRenewIndFilter = Value
        End Set
    End Property
    Public Property ExcessiveRecordsWarningInEffect() As Boolean
        Get
            Return cExcessiveRecordsWarningInEffect
        End Get
        Set(ByVal Value As Boolean)
            cExcessiveRecordsWarningInEffect = Value
        End Set
    End Property
    Public Property Sql() As String
        Get
            Return cSql
        End Get
        Set(ByVal Value As String)
            cSql = Value
        End Set
    End Property
    Public Property DateRange() As String
        Get
            Return cDateRange
        End Get
        Set(ByVal Value As String)
            cDateRange = Value
        End Set
    End Property
    Public Property InitialReportDataSuppressInEffect() As Boolean
        Get
            Return cInitialReportDataSuppressInEffect
        End Get
        Set(ByVal Value As Boolean)
            cInitialReportDataSuppressInEffect = Value
        End Set
    End Property
End Class

Public Class BulkAppointmentDataPack
    Private cDT As DataTable

    Public Class JITBulkClass
        Public Const JIT As String = "JIT"
        Public Const ApptExpress As String = "ApptExpress"
    End Class
    Public Class SaveTypeClass
        Public Const JITTrigger As String = "JIT trigger"
        Public Const JITQualify As String = "JIT qualify"
        Public Const ApptExpress As String = "Appt Express"
        Public Const ApptExpressJITTrigger As String = "Appt Express JIT trigger"
    End Class

    Public Sub New()
        cDT = New DataTable
        cDT.Columns.Add("RecordKey", GetType(System.String))
        cDT.Columns.Add("JITBulk", GetType(System.String))
        cDT.Columns.Add("SaveType", GetType(System.String))
        cDT.Columns.Add("State", GetType(System.String))
        cDT.Columns.Add("CarrierID", GetType(System.String))
        cDT.Columns.Add("Success", GetType(System.Int64))
        cDT.Columns.Add("Comments", GetType(System.String))
    End Sub

    Public Sub Init()
        cDT.Rows.Clear()
    End Sub

    Public Sub BulkAdd(ByVal SaveType As String, ByVal State As String, ByVal CarrierID As String, ByVal Success As Integer, ByVal Comments As String, Optional ByVal MakeFirstRow As Boolean = False)
        Dim dr As DataRow

        Try

            dr = cDT.NewRow
            dr("RecordKey") = State & "_" & CarrierID

            Select Case SaveType
                Case SaveTypeClass.JITTrigger, SaveTypeClass.JITQualify
                    dr("JITBulk") = JITBulkClass.JIT
                Case SaveTypeClass.ApptExpressJITTrigger, SaveTypeClass.ApptExpress
                    dr("JITBulk") = JITBulkClass.ApptExpress
            End Select

            dr("SaveType") = SaveType
            dr("State") = State
            dr("CarrierID") = CarrierID
            dr("Success") = Success
            dr("Comments") = Comments

            If MakeFirstRow Then
                cDT.Rows.InsertAt(dr, 0)
            Else
                cDT.Rows.Add(dr)
            End If

        Catch ex As Exception
            Throw New Exception("Error #803: BulkAppointmentDataPack BulkAdd. " & ex.Message)
        End Try
    End Sub

    Public Function RecordExists(ByVal State As String, ByVal CarrierID As String) As Boolean
        Dim i As Integer
        Dim Expr As String

        Try

            Expr = State & "_" & CarrierID
            For i = 0 To cDT.Rows.Count - 1
                If cDT.Rows(i)("RecordKey") = Expr Then
                    Return True
                End If
            Next
            Return False

        Catch ex As Exception
            Throw New Exception("Error #804: BulkAppointmentDataPack RecordExists. " & ex.Message)
        End Try
    End Function

    Public ReadOnly Property DT() As DataTable
        Get
            Return cDT
        End Get
    End Property
End Class