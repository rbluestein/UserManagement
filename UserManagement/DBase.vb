Imports System.Data.SqlClient

Public Class DBase
    Private cEnviro As Enviro
    Private cDictColl As New Collection

    Public Sub New()
        Dim SessionObj As System.Web.SessionState.HttpSessionState = System.Web.HttpContext.Current.Session
        cEnviro = SessionObj("Enviro")
    End Sub

    'Public Class SqlDataType
    '    Public Const BigIntSQL As String = "BigIntSQL"
    '    Public Const BitSQL As String = "BitSQL"
    '    Public Const CharSQL As String = "CharSQL"

    '    Public Const DateTimeSQL As String = "DateTimeSQL"
    '    Public Const DecimalSQL As String = "DecimalSQL"
    '    Public Const FloatSQL As String = "FloatSQL"

    '    Public Const IntSQL As String = "IntSQL"
    '    Public Const MoneySQL As String = "MoneySQL"
    '    Public Const NCharSQL As String = "NCharSQL"

    '    Public Const NumericSQL As String = "NumericSQL"
    '    Public Const NVarcharSQL As String = "NVarcharSQL"
    '    Public Const SmallDateTimeSQL As String = "SmallDateTimeSQL"

    '    Public Const SmallIntSQL As String = "SmallIntSQL"
    '    Public Const TextSQL As String = "TextSQL"
    '    Public Const TinyIntSQL As String = "TinyIntSQL"
    '    Public Const UniqueIdentifierSQL As String = "UniqueIdentifierSQL"
    '    Public Const VarcharSQL As String = "VarcharSQL"
    'End Class

    Public Enum SqlDataTypeEnum
        BitSQL = 1
        CharSQL = 2
        DateTimeSQL = 3
        IntSQL = 4
        VarcharSQL = 5
        UniqueIdentifier = 6

        'BigIntSQL = 1
        'DecimalSQL = 5
        'FloatSQL = 6
        'MoneySQL = 8
        'NCharSQL = 9
        'NumericSQL = 10
        'NVarcharSQL = 11
        'SmallDateTimeSQL = 12
        'SmallIntSQL = 13
        'TextSQL = 14
        'TinyIntSQL = 15
        'UniqueIdentifierSQL = 16
    End Enum

    Public Enum SqlDataGroupEnum
        SqlGroupBit = 1
        SqlGroupDateTime = 2
        SqlGroupFloat = 3
        SqlGroupInt = 4
        SqlGroupMoney = 5
        SqlGroupVarchar = 6
    End Enum

    Public ReadOnly Property DictColl()
        Get
            Return cDictColl
        End Get
    End Property

    'Public Function GetDTSqlDataElements(ByVal TableName As String, ByRef dt As DataTable) As DataTable
    '    Dim dt2 As DataTable
    '    dt2 = GetDTSqlDataElements(cEnviro.DefaultDatabase, TableName, dt)
    '    Return dt2
    'End Function

    'Public Function GetDTSqlDataElements(ByVal Database As String, ByVal TableName As String, ByRef dt As DataTable) As DataTable
    '    Dim i As Integer
    '    Dim Row As Integer
    '    Dim Col As Integer
    '    Dim dt2 As New DataTable
    '    Dim dr2 As DataRow
    '    Dim DictItem As DBase.DictItem

    '    For Col = 0 To dt.Columns.Count - 1
    '        dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(SqlDataElement))
    '    Next

    '    For Row = 0 To dt.Rows.Count - 1
    '        dr2 = dt2.NewRow
    '        For Col = 0 To dt.Columns.Count - 1

    '            For i = 1 To cDictColl.Count
    '                If cDictColl(i).Database = Database AndAlso cDictColl(i).Tablename = TableName AndAlso cDictColl(i).ColumnName = dt.Columns(Col).ColumnName Then
    '                    DictItem = cDictColl(i)
    '                End If
    '            Next

    '            Dim SqlData As New SqlDataElement(DictItem, dt.Rows(Row)(Col))
    '            dr2(Col) = SqlData

    '        Next
    '        dt2.Rows.Add(dr2)
    '    Next
    '    Return dt2
    'End Function

    Public Function SqlProperCase(ByVal SubjectFieldName As String, ByVal NewFieldName As String) As String
        Dim Results As String = String.Empty
        If SubjectFieldName.Length = 1 Then
            Results = Replace("UPPER(SUBSTRING(|,1, 1)) AS ", "|", SubjectFieldName)
            Results &= NewFieldName
        Else
            Results = Replace("UPPER(SUBSTRING(|,1, 1))  + LOWER(SUBSTRING(|, 2, LEN(|) -1)) AS ", "|", SubjectFieldName)
            Results &= NewFieldName
        End If
        Return Results
    End Function

    Public Sub DictionaryAddField(ByRef DictItem As DBase.DictItem)
        cDictColl.Add(DictItem)
    End Sub

    Public Sub DictionaryAddTable(ByRef Common As Common, ByVal Tablename As String)
        DictionaryAddTable(Common, cEnviro.DBHost, "UserManagement", Tablename)
    End Sub

    Public Sub DictionaryAddTable(ByRef Common As Common, ByVal DBHost As String, ByVal Database As String, ByVal Tablename As String)
        Dim i As Integer
        Dim Sql As String
        Dim ConnString As String
        Dim dt As DataTable
        Dim QueryPack As New DBase.QueryPack

        Sql = "SELECT Column_Name, Is_Nullable, Data_Type, Character_Maximum_Length, Column_Default FROM information_schema.columns WHERE table_catalog='" & Database & "' AND table_name='" & Tablename & "'"
        '' ConnString = Replace(cEnviro.ConnStringTempate, "|", Database) & DBHost
        'ConnString = cEnviro.GetConnectionString(DBHost, Database) 
        'QueryPack = ExecuteQuerySourceIsText(Tablename, Sql, ConnString, False)

        QueryPack = Common.GetDTWithQueryPack(Sql, DBHost, Database)
        dt = QueryPack.dt

        For i = 0 To dt.Rows.Count - 1
            cDictColl.Add(New DictItem(Database, Tablename, dt.Rows(i)("Column_Name"), dt.Rows(i)("Is_Nullable"), dt.Rows(i)("Data_Type"), dt.Rows(i)("Character_Maximum_Length"), dt.Rows(i)("Column_Default")))
        Next

    End Sub
    Public Class DictItem
        Private cDatabase As String
        Private cTableName As String
        Private cColumnName As String
        Private cNullable As Boolean
        Private cDataType As SqlDataTypeEnum
        Private cDataGroup As SqlDataGroupEnum
        Private cMaxLength As Integer
        Private cDefaultValue As Object

        Public Sub New(ByVal Database As String, ByVal TableName As String, ByVal ColumnName As String, ByVal Nullable As String, ByVal DataType As String, ByVal MaxLength As Object, ByVal DefaultValue As Object)
            cDatabase = Database
            cTableName = TableName
            cColumnName = ColumnName
            If Nullable.ToUpper = "YES" Then
                cNullable = True
            End If

            Select Case DataType.ToLower
                Case "bit", "char", "datetime", "int", "varchar", "uniqueidentifier"
                Case Else
                    Throw New Exception("Unsupported datatype: " & TableName & "." & ColumnName & "." & DataType)
            End Select

            Select Case DataType.ToLower
                Case "bit"
                    cDataType = SqlDataTypeEnum.BitSQL
                    cDataGroup = SqlDataGroupEnum.SqlGroupBit
                Case "char"
                    cDataType = SqlDataTypeEnum.CharSQL
                    cDataGroup = SqlDataGroupEnum.SqlGroupVarchar
                Case "datetime"
                    cDataType = SqlDataTypeEnum.DateTimeSQL
                    cDataGroup = SqlDataGroupEnum.SqlGroupDateTime
                Case "int"
                    cDataType = SqlDataTypeEnum.IntSQL
                    cDataGroup = SqlDataGroupEnum.SqlGroupInt
                Case "varchar"
                    cDataType = SqlDataTypeEnum.VarcharSQL
                    cDataGroup = SqlDataGroupEnum.SqlGroupVarchar
                Case "uniqueidentifier"
                    cDataType = SqlDataTypeEnum.UniqueIdentifier
                    cDataGroup = SqlDataGroupEnum.SqlGroupVarchar


                    'Case "bigint" : cDataType = SqlDataTypeEnum.BigIntSQL
                    'Case "decimal" : cDataType = SqlDataTypeEnum.DecimalSQL
                    'Case "float" : cDataType = SqlDataTypeEnum.FloatSQL
                    'Case "money" : cDataType = SqlDataTypeEnum.MoneySQL
                    'Case "nchar" : cDataType = SqlDataTypeEnum.NCharSQL
                    'Case "numeric" : cDataType = SqlDataTypeEnum.NumericSQL
                    'Case "nvarchar" : cDataType = SqlDataTypeEnum.NVarcharSQL
                    'Case "smalldatetime" : cDataType = SqlDataTypeEnum.SmallDateTimeSQL
                    'Case "smallint" : cDataType = SqlDataTypeEnum.SmallIntSQL
                    'Case "text" : cDataType = SqlDataTypeEnum.TextSQL
                    'Case "tinyint" : cDataType = SqlDataTypeEnum.TinyIntSQL
                    'Case "uniqueidentifier" : cDataType = SqlDataTypeEnum.UniqueIdentifierSQL
            End Select

            If Not IsDBNull(MaxLength) Then
                cMaxLength = MaxLength
            End If
            cDefaultValue = DefaultValue
        End Sub
        Public ReadOnly Property Database() As String
            Get
                Return cDatabase
            End Get
        End Property
        Public ReadOnly Property Tablename() As String
            Get
                Return cTableName
            End Get
        End Property
        Public ReadOnly Property ColumnName() As String
            Get
                Return cColumnName
            End Get
        End Property
        Public ReadOnly Property Nullable() As Boolean
            Get
                Return cNullable
            End Get
        End Property
        Public ReadOnly Property DataType() As DBase.SqlDataTypeEnum
            Get
                Return cDataType
            End Get
        End Property
        Public ReadOnly Property DataGroup() As DBase.SqlDataGroupEnum
            Get
                Return cDataGroup
            End Get
        End Property
        Public ReadOnly Property MaxLength() As Integer
            Get
                Return cMaxLength
            End Get
        End Property
        Public ReadOnly Property DefaultValue() As Object
            Get
                Return cDefaultValue
            End Get
        End Property
    End Class

    Public Class QueryPack
        Private cReturnDataTable As Boolean
        Private cReturnDataSet As Boolean
        Private cSuccess As Boolean
        Private cGenErrMsg As String
        Private cTechErrMsg As String
        Private cdt As DataTable
        Private cds As DataSet

        Public Property Success()
            Get
                Return cSuccess
            End Get
            Set(ByVal Value)
                cSuccess = Value
            End Set
        End Property

        Public ReadOnly Property GenErrMsg()
            Get
                Return GenErrMsg
            End Get
        End Property
        Public Property TechErrMsg()
            Get
                Return cTechErrMsg
            End Get
            Set(ByVal Value)
                cTechErrMsg = Value
            End Set
        End Property
        Public Property dt()
            Get
                Return cdt
            End Get
            Set(ByVal Value)
                cdt = Value
            End Set
        End Property
        Public Property ds()
            Get
                Return cds
            End Get
            Set(ByVal Value)
                cds = Value
            End Set
        End Property
    End Class

    Public Class SqlDataElement
        Private cDictItem As DBase.DictItem
        Private cValue As Object

        Public Sub New(ByRef DictItem As DBase.DictItem, ByVal Value As Object)
            cDictItem = DictItem
            cValue = Value
        End Sub

        Public Property Value()
            Get
                Return cValue
            End Get
            Set(ByVal Value)
                cValue = Value
            End Set
        End Property

        Public ReadOnly Property ToSqlF() As String
            Get
                Return cDictItem.ColumnName & "=" & GetToSqlQ()
            End Get
        End Property

        Public ReadOnly Property ToSql() As String
            Get
                Return GetToSqlQ()
            End Get
        End Property

        Public ReadOnly Property ToText()
            Get
                Select Case cDictItem.DataGroup

                    Case DBase.SqlDataGroupEnum.SqlGroupBit
                        If IsDBNull(cValue) Then
                            Return False
                        Else
                            If cValue = 1 Then
                                Return True
                            Else
                                Return False
                            End If
                        End If

                    Case DBase.SqlDataGroupEnum.SqlGroupDateTime
                        If IsDBNull(cValue) Then
                            Return String.Empty
                        Else
                            Return CType(cValue, System.DateTime)
                        End If

                    Case DBase.SqlDataGroupEnum.SqlGroupFloat
                        If IsDBNull(cValue) Then
                            Return 0
                        Else
                            Return CType(cValue, System.Double)
                        End If

                    Case DBase.SqlDataGroupEnum.SqlGroupInt
                        If IsDBNull(cValue) Then
                            Return 0
                        Else
                            Return CType(cValue, System.Int64)
                        End If

                    Case DBase.SqlDataGroupEnum.SqlGroupMoney
                        If IsDBNull(cValue) Then
                            Return FormatNumber(CType(0, System.Decimal), 2)
                        Else
                            Return FormatNumber(CType(cValue, System.Decimal), 2)
                        End If

                    Case DBase.SqlDataGroupEnum.SqlGroupVarchar
                        If IsDBNull(cValue) Then
                            Return String.Empty
                        Else
                            Return CType(cValue, System.String)
                        End If
                End Select
            End Get
        End Property

        Private Function GetToSqlQ() As String
            Select Case cDictItem.DataGroup
                Case DBase.SqlDataGroupEnum.SqlGroupDateTime, DBase.SqlDataGroupEnum.SqlGroupVarchar
                    If GetToSqlRaw() = "null" Then
                        Return "null"
                    Else
                        Return "'" & GetToSqlRaw() & "'"
                    End If
                Case Else
                    Return GetToSqlRaw()
            End Select
        End Function

        Private Function GetToSqlRaw()
            Select Case cDictItem.DataGroup

                Case DBase.SqlDataGroupEnum.SqlGroupBit
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return "0"
                            Else
                                Return cDictItem.DefaultValue
                            End If
                        End If
                    Else
                        If cValue = 0 Then
                            Return "0"
                        Else
                            Return "1"
                        End If
                    End If


                Case DBase.SqlDataGroupEnum.SqlGroupDateTime
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return "1/1/1950"
                            Else
                                Return CType(cDictItem.DefaultValue, System.String)
                            End If
                        End If
                    Else
                        Return CType(cValue, System.String)
                    End If

                Case DBase.SqlDataGroupEnum.SqlGroupFloat
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return "0"
                            Else
                                Return CType(cDictItem.DefaultValue, System.String)
                            End If
                        End If
                    Else
                        Return CType(cValue, System.String)
                    End If

                Case DBase.SqlDataGroupEnum.SqlGroupInt
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return "0"
                            Else
                                Return CType(cDictItem.DefaultValue, System.String)
                            End If
                        End If
                    Else
                        Return CType(cValue, System.String)
                    End If

                Case DBase.SqlDataGroupEnum.SqlGroupMoney
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return "0"
                            Else
                                Return CType(cDictItem.DefaultValue, System.String)
                            End If
                        End If
                    Else
                        Return CType(cValue, System.String)
                    End If


                Case DBase.SqlDataGroupEnum.SqlGroupVarchar
                    If IsDBNull(cValue) Then
                        If cDictItem.Nullable Then
                            Return "null"
                        Else
                            If IsDBNull(cDictItem.DefaultValue) Then
                                Return ""
                            Else
                                Return CType(cDictItem.DefaultValue, System.String)
                            End If
                        End If
                    Else
                        Return CType(cValue, System.String)
                    End If
            End Select
        End Function
    End Class

#Region " Formatted table "
    Public Function GetDTExtended(ByRef dt As DataTable) As DataTable
        Dim Row As Integer
        Dim Col As Integer
        Dim dt2 As New DataTable
        Dim dr2 As DataRow

        For Col = 0 To dt.Columns.Count - 1
            dt2.Columns.Add(dt.Columns(Col).ColumnName, GetType(ExtendedCell))
        Next

        For Row = 0 To dt.Rows.Count - 1
            dr2 = dt2.NewRow
            For Col = 0 To dt.Columns.Count - 1
                Select Case dt.Columns(Col).DataType.ToString.ToLower
                    Case "system.string", "system.guid"
                        Dim ExtendedCellVarchar As New ExtendedCellVarchar(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellVarchar
                        dr2(Col) = ExtendedCell
                    Case "system.int64", "system.int32", "system.int16"
                        Dim ExtendedCellInt As New ExtendedCellInt(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellInt
                        dr2(Col) = ExtendedCell
                    Case "system.boolean"
                        Dim ExtendedCellBit As New ExtendedCellBit(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellBit
                        dr2(Col) = ExtendedCell
                    Case "system.datetime"
                        Dim ExtendedCellDateTime As New ExtendedCellDateTime(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellDateTime
                        dr2(Col) = ExtendedCell
                    Case "system.double", "system.single"
                        Dim ExtendedCellFloat As New ExtendedCellFloat(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellFloat
                        dr2(Col) = ExtendedCell
                    Case "system.decimal"
                        Dim ExtendedCellMoney As New ExtendedCellMoney(dt.Rows(Row)(Col))
                        Dim ExtendedCell As ExtendedCell = ExtendedCellMoney
                        dr2(Col) = ExtendedCell
                    Case Else
                        Throw New Exception("Unprocessed data type in DBase.GetDTExtended -- " & dt.Columns(Col).DataType.ToString.ToLower & ".")
                End Select
            Next
            dt2.Rows.Add(dr2)
        Next
        Return dt2
    End Function

    Public Class ExtendedCell
        Protected cValue As Object
        Public Sub New(ByVal Value As Object)
            cValue = Value
        End Sub
    End Class

    Public Class ExtendedCellDateTime
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, DateTime)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDate(Value) Then
                    cValue = CType(Value, DateTime)
                ElseIf IsDBNull(Value) Then
                    cValue = DBNull.Value
                ElseIf Value = Nothing Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, DateTime)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    If cValue = "1/1/1950" Then
                        Return String.Empty
                    Else
                        Return CType(cValue, DateTime)
                    End If
                    'Return CType(cValue, DateTime)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                ' Return ToText
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellInt
        Inherits ExtendedCell
        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.Int64)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                ElseIf Value = Nothing Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Int64)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    Return CType(cValue, System.Int64)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellVarchar
        Inherits ExtendedCell
        Private cLength As Integer

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Length() As Integer
            Get
                Return cLength
            End Get
            Set(ByVal Value As Integer)
                cLength = Value
            End Set
        End Property
        Public Property Value()
            Get
                If IsDBNull(cValue) OrElse cValue = Nothing Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.String)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) OrElse Value = Nothing Then
                    cValue = DBNull.Value
                ElseIf IsDate(Value) Then
                    cValue = CType(Value, System.String)
                ElseIf IsNumeric(Value) Then
                    cValue = CType(Value, System.String)
                ElseIf Value = "null" Then
                    cValue = "null"
                Else
                    cValue = CType(Value, System.String)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                Dim Working As String
                If IsDBNull(cValue) OrElse cValue = Nothing Then
                    Return String.Empty
                Else
                    Working = CType(cValue, System.String)
                    ' Return Replace(Working, "~", "'")
                    Return Working
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Dim Working As String
                If ToText.length > 0 Then
                    Working = ToText
                    'Return Replace(Working, "'", "&#39;")
                    Return Replace(Working, "'", "~")
                    'Return Replace(Working, "'", "&apos;")
                    Return ToText
                Else
                    Return ToText
                End If
            End Get
        End Property
    End Class

    Public Class ExtendedCellBit
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.Int64)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    If Value Then
                        cValue = 1
                    Else
                        cValue = 0
                    End If
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    If Math.Abs(CType(cvalue, System.Int64)) = 1 Then
                        Return True
                    Else
                        Return False
                    End If
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellMoney
        Inherits ExtendedCell

        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    'Return CType(cValue, System.Decimal)
                    Return FormatNumber(CType(cValue, System.Decimal), 2)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Decimal)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    'Return CType(cValue, System.Decimal)
                    Return FormatNumber(CType(cValue, System.Decimal), 2)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class

    Public Class ExtendedCellFloat
        Inherits ExtendedCell
        Public Sub New(ByVal Value As Object)
            MyBase.New(Value)
        End Sub
        Public Property Value()
            Get
                If IsDBNull(cValue) Then
                    Return DBNull.Value
                Else
                    Return CType(cValue, System.Double)
                End If
            End Get
            Set(ByVal Value As Object)
                If IsDBNull(Value) Then
                    cValue = DBNull.Value
                Else
                    cValue = CType(Value, System.Double)
                End If
            End Set
        End Property
        Public ReadOnly Property ToText()
            Get
                If IsDBNull(cValue) Then
                    Return String.Empty
                Else
                    Return CType(cValue, System.Double)
                End If
            End Get
        End Property
        Public ReadOnly Property ToJSParm()
            Get
                Return CType(ToText, System.String)
            End Get
        End Property
    End Class
#End Region

End Class

