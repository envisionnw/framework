Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_Data
' Level:        Application module
' Version:      1.06
' Description:  data functions & procedures specific to this application
'
' Source/date:  Bonnie Campbell, 2/9/2015
' Revisions:    BLC - 2/9/2015  - 1.00 - initial version
'               BLC - 2/18/2015 - 1.01 - included subforms in fillList
'               BLC - 5/1/2015  - 1.02 - integerated into Invasives Reporting tool
'               BLC - 5/22/2015 - 1.03 - added PopulateList
'               BLC - 6/3/2015  - 1.04 - added IsUsedTargetArea
'               BLC - 12/1/2015 - 1.05 - "extra" vs target area renaming (IsUsedTargetArea > IsUsedExtraArea)
'               BLC - 6/14/2017 - 1.06 - add SetRecord(), GetRecords()
' =================================

' ---------------------------------
' SUB:          fillList
' Description:  Fill a list (or listbox like subform) from specific queries for datasheets, species or other items
' Assumptions:  Either a listbox or subform control is being populated
' Parameters:   frm - main form object
'               ctrl - either:
'                      lbx - main form listbox object (for filling a listbox control)
'                      sfrm - subform object (for populating a subform control)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC, 2/6/2015  - initial version
'   BLC, 2/18/2015 - adapted to include subform as well as listbox controls
'   BLC, 5/1/2015  - integrated into Invasives Reporting tool
' ---------------------------------
Public Sub fillList(frm As Form, ctrlSource As Control, Optional ctrlDest As Control)

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strQuery As String, strSQL As String
    
    'output to form or listbox control?
   
    'determine data source
    Select Case ctrlSource.Name
    
        Case "lbxDataSheets", "sfrmDatasheets" 'Datasheets
            strQuery = "qry_Active_Datasheets"
            strSQL = CurrentDb.QueryDefs(strQuery).SQL
            
        Case "lbxSpecies", "lbxTgtSpecies", "fsub_Species_Listbox" 'Species
            strQuery = "qry_Plant_Species"
            strSQL = CurrentDb.QueryDefs(strQuery).SQL
            
    End Select

    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)

    'set TempVars
    TempVars.Add "strSQL", strSQL

    If Not ctrlDest Is Nothing Then
        'populate list & headers
        PopulateList ctrlSource, rs, ctrlDest
    Else
        'populate only ctrlSource headers
        PopulateListHeaders ctrlSource, rs
    End If
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - fillList[mod_App_Data])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          PopulateList
' Description:  Populate listbox and similar controls from recordset
' Assumptions:  -
' Parameters:   ctrlSource - source control (listbox/listview)
'               rs - recordset used to populate control (recordset object)
'               ctrlDest - destination control (listbox/listview)
' Returns:      -
' Throws:       none
' References:   none
' Source/date:
' krish KM, Aug. 27, 2014
' http://stackoverflow.com/questions/25526904/populate-listbox-using-ado-recordset
' Adapted:      Bonnie Campbell, February 6, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/6/2015 - initial version
'   BLC - 5/10/2015 - moved to mod_List from mod_Lists
'   BLC - 5/20/2015 - changed from tbxMasterCode to tbxLUCode
'   BLC - 5/22/2015 - moved to mod_App_Data from mod_List
'   BLC - 12/1/2015 - "extra" vs. target area renaming (tbxTgtAreaID > tbxExtraAreaID, Target_Area_ID > Extra_Area_ID)
' ---------------------------------
Public Sub PopulateList(ctrlSource As Control, rs As Recordset, ctrlDest As Control)

On Error GoTo Err_Handler

    Dim frm As Form
    Dim rows As Integer, cols As Integer, i As Integer, j As Integer, matches As Integer, iZeroes As Integer
    Dim stritem As String, strColHeads As String, aryColWidths() As String

    Set frm = ctrlSource.Parent
    
    rows = rs.RecordCount
    cols = rs.Fields.Count
    
    'address no records
    If Nz(rows, 0) = 0 Then
        MsgBox "Sorry, no records found..."
        GoTo Exit_Sub
    End If
    
    'handle sfrm controls (acSubform = 112)
    If ctrlSource.ControlType = acSubform Then
        Set ctrlSource.Form.Recordset = rs
        
        ctrlSource.Form.Controls("tbxCode").ControlSource = "Code"
        ctrlSource.Form.Controls("tbxSpecies").ControlSource = "Species"
        'ctrlSource.Form.Controls("tbxMasterCode").ControlSource = "Master_PLANT_Code"
        ctrlSource.Form.Controls("tbxLUCode").ControlSource = "LUCode"
        ctrlSource.Form.Controls("tbxTransectOnly").ControlSource = "Transect_Only"
        ctrlSource.Form.Controls("tbxExtraAreaID").ControlSource = "Target_Area_ID"
        
        'set the initial record count (MoveLast to get full count, MoveFirst to set display to first)
        rs.MoveLast
        ctrlSource.Parent.Form.Controls("lblSfrmSpeciesCount").Caption = rs.RecordCount & " species"
        rs.MoveFirst
        
        GoTo Exit_Sub
    End If
    
    'fetch column widths array
    aryColWidths = Split(ctrlSource.ColumnWidths, ";")
    
    'count number of 0 width elements
    iZeroes = CountArrayValues(aryColWidths, "0")
        
    'clear out existing values
    ClearList ctrlSource
    
    'populate column names (if desired)
    If ctrlSource.ColumnHeads = True Then
        PopulateListHeaders ctrlSource, rs
        
        'populate second listbox headers if present
        If ctrlDest.ColumnHeads = True Then
            ClearList ctrlDest
            PopulateListHeaders ctrlDest, rs
        End If
    End If
    
    'populate data
    Select Case ctrlSource.RowSourceType
        Case "Table/Query"
            Set ctrlSource.Recordset = rs
        Case "Value List"
            
            'initialize
            i = 0
            
            Do Until rs.EOF
            
                'initialize item
                stritem = ""
                    
                'generate item
                For j = 0 To cols - 1
                    'check if column is displayed width > 0
                    If CInt(aryColWidths(j)) > 0 Then
                    
                        stritem = stritem & rs.Fields(j).Value & ";"
                    
                        'determine how many separators there are (";") --> should equal # cols
                        matches = (Len(stritem) - Len(Replace$(stritem, ";", ""))) / Len(";")
                        
                        'add item if not already in list --> # of ; should equal cols - 1
                        'but # in list should only be # of non-zero columns --> cols - iZeroes
                        If matches = cols - iZeroes Then
                            ctrlSource.AddItem stritem
                            'reset the string
                            stritem = ""
                        End If
                    
                    End If
                
                Next
                
                i = i + 1
                
                rs.MoveNext
            Loop
        Case "Field List"
    End Select

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          AddListToTable
' Description:  Populate table from listbox
' Assumptions:  -
' Parameters:   lbx - listbox control
' Returns:      -
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, June 3, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/3/2015 - initial version
'   BLC - 12/1/2015 - "extra" vs. target area renaming (iTgtAreaID > iExtraAreaID, Target_Area_ID > Extra_Area_ID)
' ---------------------------------
Public Sub AddListToTable(lbx As ListBox)

On Error GoTo Err_Handler

Dim aryFields() As String
Dim aryFieldTypes() As Variant
Dim strCode As String, strSpecies As String, strLUCode As String
Dim iRow As Integer, iTransectOnly As Integer, iExtraAreaID As Integer
    
    iRow = lbx.ListCount - 1 'Forms("frm_Tgt_Species").Controls("lbxTgtSpecies").ListCount - 1
    
    ReDim Preserve aryFields(0 To iRow)
        
    'header row (iRow = 0)
    aryFields(0) = "Code;Species;LUCode;Transect_Only;Extra_Area_ID"   'iRow = 0
    aryFieldTypes = Array(dbText, dbText, dbText, dbInteger, dbInteger)

    'data rows (iRow > 0)
    For iRow = 1 To lbx.ListCount - 1
        
        ' ---------------------------------------------------
        '  NOTE: listbox column MUST have a non-zero width to retrieve its value
        ' ---------------------------------------------------
         strCode = lbx.Column(0, iRow) 'column 0 = Master_PLANT_Code (Code)
         strSpecies = lbx.Column(1, iRow) 'column 1 = Species name (Species)
         strLUCode = lbx.Column(2, iRow) 'column 2 = LU_Code (LUCode)
         iTransectOnly = Nz(lbx.Column(3, iRow), 0) 'column 3 = Transect_Only (TransectOnly)
         iExtraAreaID = Nz(lbx.Column(4, iRow), 0) 'column 4 = Extra_Area_ID (ExtraAreaID)
        
        aryFields(iRow) = strCode & ";" & strSpecies & ";" & strLUCode & ";" & iTransectOnly & ";" & iExtraAreaID
        
    Next
    
    'save the existing records to temp_Listbox_Recordset & replace any existing records
    SetListRecordset lbx, True, aryFields, aryFieldTypes, "temp_Listbox_Recordset", True

Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateList[mod_App_Data])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' FUNCTION:     getParkState
' Description:  Retrieve the state associated with a park (via tlu_Parks)
' Assumptions:  Park state is properly identified in tlu_Parks
' Parameters:   parkCode - 4 character park designator
' Returns:      ParkState - 2 character state abbreviation
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 19, 2015 - for NCPN tools
' Revisions:
'   BLC - 2/19/2015  - initial version
' ---------------------------------
Public Function getParkState(ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim state As String, strSQL As String
   
    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Then
        GoTo Exit_Function
    End If
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 ParkState FROM tlu_Parks WHERE ParkCode LIKE '" & ParkCode & "';"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        state = rs.Fields("ParkState").Value
    End If
   
    'return value
    getParkState = state
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getParkState[mod_App_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     getListLastModifiedDate
' Description:  Retrieve the last modified date with a park (via tbl_Target_List)
' Assumptions:  -
' Parameters:   tgtYear - 4 digit year of list (integer)
'               parkCode - 4 character park designator (string)
' Returns:      date - last modified date (mmm-d-yyyy H:nn AMPM format) for the specified target list (string)
'                      if NULL (no last modified date) returns empty string
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 10, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/10/2015  - initial version
' ---------------------------------
Public Function getListLastModifiedDate(TgtYear As Integer, ParkCode As String) As String

On Error GoTo Err_Handler
    
    Dim strCriteria As String

    'handle only appropriate park codes
    If Len(ParkCode) <> 4 Or TgtYear < 2000 Then
        GoTo Exit_Function
    End If
    
    'set lookup criteria
    strCriteria = "Park_Code LIKE '" & ParkCode & "' AND CInt(Target_Year) = " & CInt(TgtYear)
    
    'Debug.Print strCriteria
        
    'lookup last modified date & return value
    getListLastModifiedDate = Nz(Format(DLookup("Last_Modified", "tbl_Target_List", strCriteria), "mmm-d-yyyy H:nn AMPM"), "")
    
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - getListLastModifiedDate[mod_App_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' FUNCTION:     IsUsedExtraArea
' Description:  Determine if the extra/target area is in use by a target list
' Parameters:   ExtraAreaID - extra/target area idenifier (integer)
' Returns:      boolean - true if target area is in use, false if not
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, June 3, 2015 - for NCPN tools
' Revisions:
'   BLC - 6/3/2015  - initial version
'   BLC - 12/1/2015 - "extra" vs target area renaming (IsUsedTargetArea > IsUsedExtraArea)
' ---------------------------------
Public Function IsUsedExtraArea(ExtraAreaID As Integer) As Boolean

On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim strSQL As String
    
    'default
    IsUsedExtraArea = False
    
    'generate SQL ==> NOTE: LIMIT 1; syntax not viable for Access, use SELECT TOP x instead
    strSQL = "SELECT TOP 1 Target_Area_ID FROM tbl_Target_Species WHERE Target_Area_ID = " & ExtraAreaID & ";"
            
    'fetch data
    Set db = CurrentDb
    Set rs = db.OpenRecordset(strSQL)
    
    'assume only 1 record returned
    If rs.RecordCount > 0 Then
        IsUsedExtraArea = True
    Else
        GoTo Exit_Function
    End If
       
Exit_Function:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IsUsedExtraArea[mod_App_Data])"
    End Select
    Resume Exit_Function
End Function

' ---------------------------------
' Sub:          GetRecords
' Description:  Retrieve records based on template
' Assumptions:  -
' Parameters:   Template - SQL template name (string)
' Returns:      rs - data retrieved (recordset)
' Throws:       none
' References:
'   user1938742, October 17, 2014
'   http://stackoverflow.com/questions/26422970/run-query-with-parameters-and-display-in-listbox-ms-access-2013
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/22/2016 - added templates
'   BLC - 1/9/2017 - added templates
'   BLC - 2/7/2017 - added template - s_location_with_loctypeID_sensitivity
'   BLC - 3/28/2017 - added upland templates, removed big rivers templates
'   BLC - 3/30/2017 - added option for non-parameterized queries (Else)
'   BLC - 4/3/2017 - added qc_species_by_plot_visit
' --------------------------------------------------------------------
'   BLC - 4/18/2017 - added updated version to Invasives db
' --------------------------------------------------------------------
'   BLC - 4/18/2017 - adjusted for invasives templates
'   BLC - 4/24/2017 - added microhabitat surface & species templates
' ---------------------------------
Public Function GetRecords(Template As String) As DAO.Recordset
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim rs As DAO.Recordset
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(Template)
        
            Select Case Template
                        
                Case "s_access_level"
                    '-- required parameters --
                    .Parameters("lvl") = TempVars("tempLvl")
                    
                    'clear the tempvar
                    TempVars.Remove "tempLvl"
                                                                                                               
                Case "s_get_parks"
                    '-- required parameters --
                                                                    
                Case "s_park_id"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
           
                Case "s_template_num_records"
                    '-- required parameters --

                Case "qc_ndc_notrecorded_all_methods_by_plot_visit", _
                    "qc_photos_missing_by_plot_visit", _
                    "qc_species_by_plot_visit"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("pid") = TempVars("plotID")
                    .Parameters("vdate") = TempVars("SampleDate")
                
                Case "s_tsys_datasheet_defaults"
                    '-- required parameters --
                
                Case "s_surface"
                    '-- required parameters --
                
                Case "s_surface_by_ID"
                    '-- required parameters --
                    .Parameters("sid") = TempVars("SurfaceID")
                
                Case "s_speciescover_by_transect"
                    '-- required parameters --
                    .Parameters("pkcode") = TempVars("ParkCode")
                    .Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")
                
                Case "s_surfacecover_by_transect"
                    '-- required parameters --
                    '.Parameters("pkcode") = TempVars("ParkCode")
                    '.Parameters("eid") = TempVars("Event_ID")
                    .Parameters("tid") = TempVars("Transect_ID")
                
                Case Else
                    'handle other non-parameterized queries
                    
            End Select
            
            Set rs = .OpenRecordset(dbOpenDynaset)
            
        End With
        
    End With
    
    Set GetRecords = rs
    
Exit_Handler:
    Exit Function
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetRecords[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' Function:     SetRecord
' Description:  Insert/update/delete record based on template
' Assumptions:  -
' Parameters:   template - SQL template name (string)
'               params - array of parameters for template (variant)
' Returns:      id - ID of record inserted, updated, deleted (long integer)
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, July 26, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 7/26/2016 - initial version
'   BLC - 9/21/2016 - updated i_login parameters
'   BLC - 10/24/2016 - added flag templates (contact, site, mod wentworth)
'   BLC - 10/28/2016 - updated TempVars("ContactID") -> TempVars("AppUserID"), updated i_task
'   BLC - 1/24/2017 - added IsNPS flag parameter for contacts
'   BLC - 3/24/2017 - set SkipRecordAction = False for uplands, removed unused big rivers cases,
'                     added uplands cases, delete cases
'   BLC - 3/29/2017 - added FieldOK, FieldCheck, Dependencies parameters for templates
'   BLC - 4/24/2017 - add surface/species cover, set SkipRecordAction = false (invasives, uplands)
' ---------------------------------
Public Function SetRecord(Template As String, Params As Variant) As Long
On Error GoTo Err_Handler
    
    Dim db As DAO.Database
    Dim qdf As DAO.QueryDef
    Dim SkipRecordAction As Boolean
    Dim ID As Long
    
    'exit w/o values
    If Not IsArray(Params) Then GoTo Exit_Handler
    
    'default <-- upland/invasives donot have RecordAction table implemented so skip!
    SkipRecordAction = True 'False
            
    'default ID (if not set as param)
    ID = 0
    
    Set db = CurrentDb
    
    With db
        Set qdf = .QueryDefs("usys_temp_qdf")
        
        With qdf
        
            'check if record exists in site
            .SQL = GetTemplate(Template)
            
            '-------------------
            ' set SQL parameters --> .Parameters("") = params()
            '-------------------
            
            '-------------------------------------------------------------------------
            ' NOTE:
            '   param(0) --> reserved for record action RefTable (ReferenceType)
            '   last param(x) --> used as record ID for updates
            '-------------------------------------------------------------------------
            Select Case Template
            
        '-----------------------
        '  INSERTS
        '-----------------------
                Case "i_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)  'record ID
                    .Parameters("num") = Params(2)  'number of records
                    .Parameters("fok") = Params(3)  'field ok? (QC pass/fail)
                    
                Case "i_template"
                    '-- required parameters --
                    .Parameters("tname") = Params(1)        'TemplateName
                    .Parameters("contxt") = Params(2)       'Context
                    '.Parameters("tmpl").Type = dbMemo       'set it to a memo field
                    'Limit template SQL to 255 characters to avoid
                    'error 3271 SetRecord mod_App_Data Invalid property value.
                    'templates > 255 characters must be edited directly in the table
                    .Parameters("tmpl") = Left(Params(3), 255) 'TemplateSQL
                    .Parameters("rmks") = Params(4)         'Remarks
                    .Parameters("effdate") = Params(5)      'EffectiveDate
                    .Parameters("cid") = Params(6)          'CreatedBy_ID (contactID)
                    .Parameters("prms") = Params(7)         'Params
                    .Parameters("syntx") = Params(8)        'Syntax
                    .Parameters("vers") = Params(9)         'Version
                    .Parameters("sflag") = Params(10)       'IsSupported
                    .Parameters("lmid") = TempVars("AppUserID") 'lastmodifiedID
                    .Parameters("fqc") = Params(11)         'FieldCheck
                    .Parameters("fok") = Params(12)         'FieldOK
                    .Parameters("dep") = Params(13)         'Dependencies
                
                Case "i_surface_cover"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)
                    .Parameters("sid") = Params(2)
                    .Parameters("pct") = Params(3)
                    
'                    .Parameters("") = Params(1)
'                    .Parameters("") = Params(2)
'                    .Parameters("") = Params(3)
                
        '-----------------------
        '  UPDATES
        '-----------------------
                Case "u_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
                    .Parameters("num") = Params(2)
                    .Parameters("fok") = Params(3)
                    
                Case "u_template"
                    '-- required parameters --
                    .Parameters("id") = Params(1)
                
                Case "u_surface_cover"
                    '-- required parameters --
                    .Parameters("qid") = Params(1)
                    .Parameters("sid") = Params(2)
                    .Parameters("pct") = Params(3)
                    '.Parameters("sfcid") = Params(4)
                    
        '-----------------------
        '  DELETES
        '-----------------------
                Case "d_num_records_all"
                    '-- required parameters --
                
                Case "d_num_records"
                    '-- required parameters --
                    .Parameters("rid") = Params(1)
            
            End Select
            
            .Execute dbFailOnError
                
    ' -------------------
    '  Record Action
    ' -------------------
            'handle unrecorded actions & those which don't generate an ID
            If SkipRecordAction Then GoTo Exit_Handler
            
            If ID = 0 Then
                'retrieve identity
                ID = db.OpenRecordset("SELECT @@IDENTITY;")(0)
            End If
            
            'set record action
            .SQL = GetTemplate("i_record_action")
                                            
            '-- required parameters --
            .Parameters("RefTable") = Params(0)
            .Parameters("RefID") = ID
            .Parameters("ID") = TempVars("AppUserID") 'TempVars("ContactID")
            .Parameters("Activity") = "DE"
            .Parameters("ActionDate") = CDate(Format(Now(), "YYYY-mm-dd hh:nn:ss AMPM"))
                                
            .Execute dbFailOnError
            
            'cleanup
            .Close
        
        End With

        SetRecord = ID
    End With
                
Exit_Handler:
    'cleanup
    Set qdf = Nothing
    Set db = Nothing

    Exit Function
Err_Handler:
    Select Case Err.Number

      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetRecord[mod_App_Data])"
    End Select
    Resume Exit_Handler
End Function


' ---------------------------------
' SUB:          CollapseRows
' Description:  Collapses TCount, PctCover, SE for one species/IsDead into one row
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
'   R.Hicks, Sept 15, 2002
'   http://www.utteraccess.com/forum/copy-table-structure-vb-t117555.html
' Source/date:  Bonnie Campbell, June 22 2017
' Adapted:      -
' Revisions:    BLC - 6/22/2017 - initial version
' ---------------------------------
Public Sub CollapseRows(tbl As String) 'tdf As DAO.TableDef)
On Error GoTo Err_Handler

    Dim db As DAO.Database
    Dim tdf As DAO.TableDef
    Dim col As field
    Dim rs As DAO.Recordset
    Dim rsPctCover As DAO.Recordset
    Dim rsSE As DAO.Recordset
    Dim strTableNew As String
    Dim strCol As String
    Dim Park As String, VisitYear As String, Species As String, CommonName As String, _
        IsDead As String, Route As String
    Dim PrevPark As String, PrevVisitYear As String, PrevSpecies As String, _
        PrevCommonName As String, PrevIsDead As String, PrevRoute As String
    Dim TCount As Integer
    Dim PctCover As String
    Dim SE As Double
    Dim Concat As String, PrevConcat As String
    
    strTableNew = tbl & "_NEW"
    
    PrevPark = ""
    PrevVisitYear = ""
    PrevSpecies = ""
    PrevCommonName = ""
    PrevIsDead = ""
    
    Set db = CurrentDb
    Set tdf = db.TableDefs(tbl)
    
    Set rs = db.OpenRecordset(tdf.Name)
    
    'remove table if it exists
    If TableExists(strTableNew) Then DoCmd.DeleteObject acTable, strTableNew
    
    'create empty table w/ same columns to fill into
    DoCmd.TransferDatabase acExport, "Microsoft Access", db.Name, acTable, tdf.Name, strTableNew, True
    
    With tdf
        'iterate through ALL
        Do Until rs.BOF And rs.EOF
        
            'iterate result columns (fields)
            For Each col In tdf.Fields
            
                'get park, visit year, species, common name & isdead
                If col.OrdinalPosition = 1 Then Park = rs!Unit_Code
                If col.OrdinalPosition = 2 Then VisitYear = rs!Visit_Year
                If col.OrdinalPosition = 3 Then Species = rs!Species
                If col.OrdinalPosition = 4 Then CommonName = rs!Master_Common_Name
                If col.OrdinalPosition = 5 Then IsDead = rs!IsDead
            
                'ignore 1-5 (static Park, Year, Species, Master Common Name, IsDead)
                If col.OrdinalPosition > 5 Then
                
                    'get column & route name
                    strCol = col.Name
                    Route = Left(strCol, InStr(col.Name, ") ") + 1)
                                     
                    Select Case Replace(strCol, Route, "")
                        Case "TCount"
                            TCount = col.Value
                        Case "AvgCover"
                            PctCover = col.Value
                        Case "SE"
                            SE = col.Value
                    End Select

                    'concatenate for comparison
                    Concat = Park & VisitYear & Species & CommonName & IsDead & Route
                    
                    If Concat <> PrevConcat Then
                
                    'add these to the new table if they don't already exist
                    End If
    
                    If PrevSpecies = rs!Species And PrevIsDead = rs!IsDead Then
                    
                    End If
                End If
                
                'capture the previous values
                PrevConcat = Concat
                PrevPark = Park
                PrevVisitYear = VisitYear
                PrevSpecies = Species
                PrevCommonName = CommonName
                PrevIsDead = IsDead
                
           Next
           
        Loop
    
    End With

Exit_Procedure:
    Set tdf = Nothing
    Set rs = Nothing
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CollapseRows[frm_Species_Cover_by_Route])"
    End Select
    Resume Exit_Procedure
End Sub