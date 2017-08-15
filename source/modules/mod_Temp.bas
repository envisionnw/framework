Option Compare Database
Option Explicit

Public Sub Delete_All_Records()
Dim strSQL As String
Dim strTables(11) As String
Const cstrSQL As String = "DELETE * FROM "
Dim i As Integer

strTables(0) = "tbl_Db_Revisions"
strTables(1) = "tbl_Db_Meta"
strTables(2) = "tbl_Event_Details"
strTables(3) = "tbl_Field_Data"
strTables(4) = "tbl_Data_Locations"
strTables(5) = "xref_Event_Contacts"
strTables(6) = "tlu_Contacts"
strTables(7) = "tbl_Events"
strTables(8) = "tbl_Event_Group"
strTables(9) = "tbl_Locations"
strTables(10) = "tbl_Sites"

For i = 0 To UBound(strTables) - 1
    strSQL = cstrSQL & strTables(i)
    CurrentDb.Execute strSQL
Next i

End Sub