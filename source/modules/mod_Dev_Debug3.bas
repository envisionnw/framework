Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Dev_Debug
' Level:        Development module
' Version:      1.02
'
' Description:  Debugging related functions & procedures for version control
'
' Source/date:  Bonnie Campbell, 2/12/2015
' Revisions:    BLC - 5/27/2015 - 1.00 - initial version
'               BLC - 7/7/2015  - 1.01 - added GetErrorTrappingOption()
'               BLC - 4/6/2017 - 1.02 - added SearchQueries(), SearchDB()
' =================================

' ===================================================================================
'  NOTE:
'       Functions and subroutines within this module are for debugging and test
'       purposes.
'
'       When the application is ready for release, this module can be
'       removed without negative impact to the application.
'
'       All mod_Debug_XX (debugging) and VCS_XX (version control system) modules can also be removed.
' ===================================================================================

' ---------------------------------
' SUB:          ChangeMSysConnection
' Description:  Change connection value for a table w/in MSys_Objects (which cannot/shouldn't be directly edited)
' Assumptions:  -
' Parameters:   strTable - table name to change (string)
'               strConn - new connection string (e.g. "ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
' Joe Kendall, 8/25/2003
' http://www.experts-exchange.com/Database/MS_Access/Q_20615117.html
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeMSysConnection(ByVal strTable As String, ByVal strConn As String)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTable)

    'Change the connect value
    tdf.connect = strConn '"ODBC;DATABASE=pubs;UID=sa;PWD=;DSN=Publishers"
    
Exit_Sub:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeMSysConnection[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ChangeMSysDb
' Description:  Change database value for a table w/in MSys_Objects (which cannot/shouldn't be directly edited)
' Assumptions:  -
' Parameters:   strTable - table name to change (string)
'               strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:
' Joe Kendall, 8/25/2003
' http://www.experts-exchange.com/Database/MS_Access/Q_20615117.html
' Adapted:      Bonnie Campbell, May 27, 2015 - for NCPN tools
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeMSysDb(ByVal strTable As String, ByVal strDbPath As String)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef

    Set db = CurrentDb()
    Set tdf = db.TableDefs(strTable)

    'Change the database value
    tdf.connect = ";DATABASE=" & strDbPath
    
    tdf.RefreshLink
    
Exit_Sub:
    Set tdf = Nothing
    db.Close
    Set db = Nothing
    
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeMSysDb[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          ChangeTSysDb
' Description:  Change database value for a table w/in tsys_Link_Files & tsys_Link_Dbs
' Assumptions:  Tables (tsys_Link_Files & tsys_Link_Dbs) exist with fields as noted
' Parameters:   strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub ChangeTSysDb(ByVal strDbPath As String)
On Error GoTo Err_Handler
    
    Dim strDbFile As String, strSQL As String
    
    'get db file name
    strDbFile = ParseFileName(strDbPath)
    
    DoCmd.SetWarnings False
    
    'update tsys_Link_Files
    strSQL = "UPDATE tsys_Link_Files SET Link_file_path = '" & strDbPath & "' WHERE Link_file_name = '" & strDbFile & "';"
    DoCmd.RunSQL (strSQL)
    
   'update tsys_Link_Dbs
    strSQL = "UPDATE tsys_Link_Dbs SET File_path = '" & strDbPath & "' WHERE Link_db = '" & strDbFile & "';"
    DoCmd.RunSQL (strSQL)
    
    DoCmd.SetWarnings True

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - ChangeTSysDb[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          SetDebugDbPaths
' Description:  Change database paths for debugging in MSys_Objects, tsys_Link_Files, & tsys_Link_Dbs
' Assumptions:  Tables (tsys_Link_Files & tsys_Link_Dbs) exist with fields as noted
'               tsys_Link_Tables exists and includes desired tables
' Parameters:   strDbPath - new database path (string) (e.g. "C:\__TEST_DATA\mydb.accdb")
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub SetDebugDbPaths(ByVal strDbPath As String)
On Error GoTo Err_Handler
    Dim rs As DAO.Recordset
    Dim strDb As String, strTable As String
    
    'change the tsys_Link_Files & tsys_Link_Dbs tables
    ChangeTSysDb strDbPath
    
    'get db name
    strDb = ParseFileName(strDbPath)
    
    'iterate through linked tables w/in tsys_Link_Tables
    Set rs = CurrentDb.OpenRecordset("tsys_Link_Tables", dbOpenDynaset)
    
    If Not (rs.BOF And rs.EOF) Then
    
        Do Until rs.EOF
            
            'match table source db
            If rs!Link_db = strDb Then
                
                strTable = rs!Link_table
                ChangeMSysDb strTable, strDbPath
            
            End If
        
            rs.MoveNext
        Loop
        
    End If
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetDebugDbPaths[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          DebugTest
' Description:  Run debug testing routines as noted within the subroutine.
' Assumptions:  This subroutine will be modified as needed during testing.
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, May 27, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 5/27/2015 - initial version
' ---------------------------------
Public Sub DebugTest()
On Error GoTo Err_Handler

    Dim strDbPath As String, strDb As String
    Dim i As Integer
    

    'invasives be
'    strDbPath = "C:\___TEST_DATA\test\Invasives_be.accdb"
    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\Invasives_be.accdb"
    strDb = ParseFileName(strDbPath)
    
    SetDebugDbPaths strDbPath
    
    'NCPN master plants
'    strDbPath = "C:\___TEST_DATA\NCPN_Master_Species.accdb"
    strDbPath = "Z:\_____LIB\dev\git_projects\TEST_DATA\test2\NCPN_Master_Species.accdb"
    strDb = ParseFileName(strDbPath)

    SetDebugDbPaths strDbPath


    'progress bar test
    DoCmd.OpenForm "frm_ProgressBar", acNormal
    
    For i = 1 To 10
        
        Forms("frm_ProgressBar").Increment i * 10, "Preparing report..."
    Next

    'test parsing
    ParseFileName ("C:\___TEST_DATA\test_BE_new\Invasives_be.accdb")


Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DebugTest[mod_Dev_Debug])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     GetErrorTrappingOption
' Description:  Determine the error trapping option setting.
' Assumptions:  -
' Parameters:   -
' Returns:      String representing the IDE's error trapping setting.
' Throws:       none
' References:   -
' Source/date:  Luke Chung, date unknown
'               http://www.fmsinc.com/tpapers/vbacode/debug.asp
' Adapted:      Bonnie Campbell, July 7, 2015 - for NCPN tools
' Revisions:
'   BLC - 7/7/2015 - initial version
' ---------------------------------
Function GetErrorTrappingOption() As String
On Error GoTo Err_Handler

  Dim strSetting As String
  
  Select Case Application.GetOption("Error Trapping")
    Case 0
      strSetting = "Break on All Errors"
    Case 1
      strSetting = "Break in Class Modules"
    Case 2
      strSetting = "Break on Unhandled Errors"
  End Select
  GetErrorTrappingOption = strSetting

Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetErrorTrappingOption[mod_Dev_Debug])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     SearchQueries
' Description:  Determine which queries contain a certain text value
' Assumptions:  -
' Parameters:   SearchText - text to find in the query (string)
'               ShowSQL - show the query SQL (boolean, false = default)
'               QryName - query to search (string, default = * which includes all queries)
' Returns:      Debug.Print's the name of each query that contains the text.
' Throws:       none
' References:   -
' Source/date:  mwolfe02, October 20, 2011
'               http://stackoverflow.com/questions/7831071/how-to-find-all-queries-related-to-table-in-ms-access
' Adapted:      Bonnie Campbell, April 6, 2017 - for NCPN tools
' Revisions:
'   BLC - 4/6/2017 - initial version
' ---------------------------------
Sub SearchQueries(SearchText As String, _
                  Optional ShowSQL As Boolean = False, _
                  Optional QryName As String = "*")
On Error GoTo Err_Handler
    
    Dim QDef As QueryDef

    For Each QDef In CurrentDb.QueryDefs
        If QDef.Name Like QryName Then
            If InStr(QDef.SQL, SearchText) > 0 Then
                Debug.Print QDef.Name
                If ShowSQL Then Debug.Print QDef.SQL & vbCrLf
            End If
        End If
    Next QDef

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SearchQueries[mod_Dev_Debug])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' FUNCTION:     SearchDB
' Description:  Determine which objects contain a certain text value
' Assumptions:  -
' Parameters:   SearchText - text to find (string)
'               ObjType - type(s) of objects to check (AcObjectType, acDefault = default)
'               ObjName - name of object to search (string, default = * which includes all objects)
' Returns:      Debug.Print's the name of each object that contains the text.
' Throws:       none
' References:   -
' Source/date:  mwolfe02, October 20, 2011
'               http://stackoverflow.com/questions/7831071/how-to-find-all-queries-related-to-table-in-ms-access
' Adapted:      Bonnie Campbell, April 6, 2017 - for NCPN tools
' Revisions:
'   MW  - 10/20/200x - initial version
'   MW  - 1/19/2009  - limited search by object name pattern
'   BLC - 4/6/2017   - initial version for NCPN tools, added casing
' ---------------------------------
Sub SearchDB(SearchText As String, _
             Optional ObjType As AcObjectType = acDefault, _
             Optional ObjName As String = "*")
On Error GoTo Err_Handler

    Dim db As Database, obj As AccessObject, ctl As Control, prop As Property
    Dim frm As Form, rpt As Report, mdl As Module
    Dim oLoaded As Boolean, found As Boolean, instances As Long
    Dim sline As Long, scol As Long, eline As Long, ecol As Long
    Dim ary() As Variant, oType As Variant

    Set db = CurrentDb
    Application.Echo False

    'set array
    If acDefault Then
        'do for all
        ary = Array(acQuery, acForm, acMacro, acModule, acReport)
    Else
        ary = Array(ObjType)
    End If
    
    'iterate
    For Each oType In ary
        
        'search object types
        Select Case oType
            '------- Queries ----------
            Case acQuery
                Debug.Print "Queries:"
                SearchQueries SearchText, False, ObjName
                Debug.Print vbCrLf
    
            '------- Forms ----------
            Case acForm
                Debug.Print "Forms:"
                On Error Resume Next
                For Each obj In CurrentProject.AllForms
                    If obj.Name Like ObjName Then
                        oLoaded = obj.IsLoaded
                        If Not oLoaded Then DoCmd.OpenForm obj.Name, acDesign, , , , acHidden
                        Set frm = Application.Forms(obj.Name)
                        For Each prop In frm.Properties
                            Err.Clear
                            If InStr(prop.Value, SearchText) > 0 Then
                                If Err.Number = 0 Then
                                    Debug.Print "Form: " & frm.Name & _
                                                "  Property: " & prop.Name & _
                                                "  Value: " & prop.Value
                                End If
                            End If
                        Next prop
                        If frm.HasModule Then
                            sline = 0: scol = 0: eline = 0: ecol = 0: instances = 0
                            found = frm.Module.Find(SearchText, sline, scol, eline, ecol)
                            Do Until Not found
                                instances = instances + 1
                                sline = eline + 1: scol = 0: eline = 0: ecol = 0
                                found = frm.Module.Find(SearchText, sline, scol, eline, ecol)
                            Loop
                            If instances > 0 Then Debug.Print "Form: " & frm.Name & _
                               "  Module: " & instances & " instances"
        
                        End If
                        For Each ctl In frm.Controls
                            For Each prop In ctl.Properties
                                Err.Clear
                                If InStr(prop.Value, SearchText) > 0 Then
                                    If Err.Number = 0 Then
                                        Debug.Print "Form: " & frm.Name & _
                                                    "  Control: " & ctl.Name & _
                                                    "  Property: " & prop.Name & _
                                                    "  Value: " & prop.Value
                                    End If
                                End If
                            Next prop
                        Next ctl
                        Set frm = Nothing
                        If Not oLoaded Then DoCmd.Close acForm, obj.Name, acSaveNo
                        DoEvents
                    End If
                Next obj
                On Error GoTo Err_Handler
                Debug.Print vbCrLf
    
            '------- Modules ----------
            Case acModule
                Debug.Print "Modules:"
                For Each obj In CurrentProject.AllModules
                    If obj.Name Like ObjName Then
                        oLoaded = obj.IsLoaded
                        If Not oLoaded Then DoCmd.OpenModule obj.Name
                        Set mdl = Application.Modules(obj.Name)
                        sline = 0: scol = 0: eline = 0: ecol = 0: instances = 0
                        found = mdl.Find(SearchText, sline, scol, eline, ecol)
                        Do Until Not found
                            instances = instances + 1
                            sline = eline + 1: scol = 0: eline = 0: ecol = 0
                            found = mdl.Find(SearchText, sline, scol, eline, ecol)
                        Loop
                        If instances > 0 Then Debug.Print obj.Name & ": " & instances & " instances"
                        Set mdl = Nothing
                        If Not oLoaded Then DoCmd.Close acModule, obj.Name
                    End If
                Next obj
                Debug.Print vbCrLf
    
            '------- Macros ----------
            Case acMacro
                'Debug.Print "Macros:"
                'Debug.Print vbCrLf
    
            '------- Reports ----------
            Case acReport
                Debug.Print "Reports:"
                On Error Resume Next
                For Each obj In CurrentProject.AllReports
                    If obj.Name Like ObjName Then
                        oLoaded = obj.IsLoaded
                        If Not oLoaded Then DoCmd.OpenReport obj.Name, acDesign
                        Set rpt = Application.Reports(obj.Name)
                        For Each prop In rpt.Properties
                            Err.Clear
                            If InStr(prop.Value, SearchText) > 0 Then
                                If Err.Number = 0 Then
                                    Debug.Print "Report: " & rpt.Name & _
                                                "  Property: " & prop.Name & _
                                                "  Value: " & prop.Value
                                End If
                            End If
                        Next prop
                        If rpt.HasModule Then
                            sline = 0: scol = 0: eline = 0: ecol = 0: instances = 0
                            found = rpt.Module.Find(SearchText, sline, scol, eline, ecol)
                            Do Until Not found
                                instances = instances + 1
                                sline = eline + 1: scol = 0: eline = 0: ecol = 0
                                found = rpt.Module.Find(SearchText, sline, scol, eline, ecol)
                            Loop
                            If instances > 0 Then Debug.Print "Report: " & rpt.Name & _
                               "  Module: " & instances & " instances"
        
                        End If
                        For Each ctl In rpt.Controls
                            For Each prop In ctl.Properties
                                If InStr(prop.Value, SearchText) > 0 Then
                                    Debug.Print "Report: " & rpt.Name & _
                                                "  Control: " & ctl.Name & _
                                                "  Property: " & prop.Name & _
                                                "  Value: " & prop.Value
                                End If
                            Next prop
                        Next ctl
                        Set rpt = Nothing
                        If Not oLoaded Then DoCmd.Close acReport, obj.Name, acSaveNo
                        DoEvents
                    End If
                Next obj
                On Error GoTo Err_Handler
                Debug.Print vbCrLf
        
        End Select
    
    Next

Exit_Handler:
    Application.Echo True
    Exit Sub
    
Err_Handler:
    Application.Echo True
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SearchDB[mod_Dev_Debug])"
    End Select
    Debug.Assert False
    Resume Exit_Handler
End Sub

Public Sub runtest()
    'SearchDB "tbl_Quadrat_Species"
    'SearchDB "tbl_Quadrat_Transect"
    'SearchDB "qry_Transect_Select"
'     Dim qdf As QueryDef
'    Set qdf = CurrentDb.QueryDefs("Query6")
'
'    'save original SQL
'    Debug.Print qdf.SQL
    
    
    Dim tbl As String 'DAO.TableDef

    'Set tbl = CurrentDb.TableDefs("DINO_2014_SpeciesCover_by_Route_Result")
    tbl = "DINO_2014_SpeciesCover_by_Route_Result"
    CollapseRows tbl

End Sub