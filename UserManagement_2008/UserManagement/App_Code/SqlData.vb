'Public Class SqlData
'    Private cDictItem As DBase.DictItem
'    Private cValue As Object

'    Public Sub New(ByRef DictItem As DBase.DictItem, ByVal Value As Object)
'        cDictItem = DictItem
'    End Sub

'    Public Property Value()
'        Get
'            Return cValue
'        End Get
'        Set(ByVal Value)
'            cValue = Value
'        End Set
'    End Property

'    Public ReadOnly Property ToSql() As String
'        Get
'            Return GetToSql()
'        End Get
'    End Property

'    Public ReadOnly Property ToSqlQ() As String
'        Get
'            Select Case cDictItem.DataGroup
'                Case DBase.SqlDataGroupEnum.SqlGroupDateTime, DBase.SqlDataGroupEnum.SqlGroupVarchar
'                    If ToSql = "null" Then
'                        Return "null"
'                    Else
'                        Return "'" & ToSql & "'"
'                    End If
'                Case Else
'                    Return ToSql
'            End Select
'        End Get
'    End Property

'    Public ReadOnly Property ToSqlF() As String
'        Get
'            Return cDictItem.ColumnName & "=" & ToSqlQ
'        End Get
'    End Property

'    Private Function GetToSql()
'        Select Case cDictItem.DataGroup

'            Case DBase.SqlDataGroupEnum.SqlGroupBit
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return "0"
'                        Else
'                            Return cDictItem.DefaultValue
'                        End If
'                    End If
'                Else
'                    If cValue = 0 Then
'                        Return "0"
'                    Else
'                        Return "1"
'                    End If
'                End If


'            Case DBase.SqlDataGroupEnum.SqlGroupDateTime
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return "1/1/1950"
'                        Else
'                            Return CType(cDictItem.DefaultValue, System.String)
'                        End If
'                    End If
'                Else
'                    Return CType(cValue, System.String)
'                End If

'            Case DBase.SqlDataGroupEnum.SqlGroupFloat
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return "0"
'                        Else
'                            Return CType(cDictItem.DefaultValue, System.String)
'                        End If
'                    End If
'                Else
'                    Return CType(cValue, System.String)
'                End If

'            Case DBase.SqlDataGroupEnum.SqlGroupInt
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return "0"
'                        Else
'                            Return CType(cDictItem.DefaultValue, System.String)
'                        End If
'                    End If
'                Else
'                    Return CType(cValue, System.String)
'                End If

'            Case DBase.SqlDataGroupEnum.SqlGroupMoney
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return "0"
'                        Else
'                            Return CType(cDictItem.DefaultValue, System.String)
'                        End If
'                    End If
'                Else
'                    Return CType(cValue, System.String)
'                End If


'            Case DBase.SqlDataGroupEnum.SqlGroupVarchar
'                If IsDBNull(cValue) Then
'                    If cDictItem.Nullable Then
'                        Return "null"
'                    Else
'                        If IsDBNull(cDictItem.DefaultValue) Then
'                            Return ""
'                        Else
'                            Return CType(cDictItem.DefaultValue, System.String)
'                        End If
'                    End If
'                Else
'                    Return CType(cValue, System.String)
'                End If
'        End Select
'    End Function

'    Public ReadOnly Property ToApp()
'        Get
'            Select Case cDictItem.DataGroup

'                Case DBase.SqlDataGroupEnum.SqlGroupBit
'                    If IsDBNull(cValue) Then
'                        Return False
'                    Else
'                        If cValue = 1 Then
'                            Return True
'                        Else
'                            Return False
'                        End If
'                    End If

'                Case DBase.SqlDataGroupEnum.SqlGroupDateTime
'                    If IsDBNull(cValue) Then
'                        Return String.Empty
'                    Else
'                        Return CType(cValue, System.DateTime)
'                    End If

'                Case DBase.SqlDataGroupEnum.SqlGroupFloat
'                    If IsDBNull(cValue) Then
'                        Return 0
'                    Else
'                        Return CType(cValue, System.Double)
'                    End If

'                Case DBase.SqlDataGroupEnum.SqlGroupInt
'                    If IsDBNull(cValue) Then
'                        Return 0
'                    Else
'                        Return CType(cValue, System.Int64)
'                    End If

'                Case DBase.SqlDataGroupEnum.SqlGroupMoney
'                    If IsDBNull(cValue) Then
'                        Return FormatNumber(CType(0, System.Decimal), 2)
'                    Else
'                        Return FormatNumber(CType(cValue, System.Decimal), 2)
'                    End If

'                Case DBase.SqlDataGroupEnum.SqlGroupVarchar
'                    If IsDBNull(cValue) Then
'                        Return String.Empty
'                    Else
'                        Return CType(cValue, System.String)
'                    End If
'            End Select
'        End Get
'    End Property
'End Class


