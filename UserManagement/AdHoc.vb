Public Class AdHoc

    ''''Public Sub InsertRecID(ByVal Enviro As Enviro, ByVal Common As Common)
    ''''    Dim i As Integer
    ''''    Dim dt As DataTable
    ''''    Dim RecID As Integer
    ''''    Dim Querypack As DBase.QueryPack
    ''''    Dim Sql As String

    ''''    dt = Common.GetDT("Select UserID from Users")
    ''''    For i = 0 To dt.Rows.Count - 1

    ''''        RecID = Common.GetNewRecordID("Users", "RecID")

    ''''        Sql = "UPDATE UserManagement..Users SET RecID = " & RecID.ToString & " WHERE UserID = '" & dt.Rows(i)("UserID") & "'"
    ''''        System.Diagnostics.Debug.WriteLine(i.ToString & "  " & dt.Rows(i)("UserID"))
    ''''        Querypack = Common.ExecuteNonQueryWithQuerypack(Sql)
    ''''        If Not Querypack.Success Then
    ''''            Stop
    ''''        End If

    ''''    Next

    ''''End Sub
End Class
