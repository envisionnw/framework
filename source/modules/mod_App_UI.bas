Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_App_UI
' Level:        Application module
' Version:      1.15
' Description:  Application User Interface related functions & subroutines
'
' Source/date:  Bonnie Campbell, April 2015
' Revisions:    BLC, 4/30/2015 - 1.00 - initial version
'               ----------- invasives reports -----------------
'               BLC, 5/26/2015 - 1.01 - added PopulateSpeciesPriorities function from mod_Species
'               BLC, 6/1/2015  - 1.02 - changed View to Search tab
'               BLC, 6/12/2015 - 1.03 - added EnableTargetTool button
'               ----------- big rivers ------------------------
'               BLC, 6/30/2015 - 1.04 - added ClearFields()
'               BLC, 7/27/2015 - 1.05 - added SetHints()
'               ----------- uplands ---------------------------
'               BLC, 8/21/2015 - 1.06 - added CaptureEscapeKey
'               BLC, 2/3/2016  - 1.07 - added SetNoDataCheckbox()
'               BLC, 2/9/2016  - 1.08 - added public dictionary for NoData checkboxes
'                                       dictionary is used within subforms to identify if checkboxes
'                                       should be checked, GetNoDataCollected(), SetNoDataCollected()
'               BLC, 2/9/2016 - 1.09 - added constants, functions & subroutine supporting transect overlays
'                                       (LWA_ALPHA, GWL_EXSTYLE, WS_EX_LAYERED, GetWindowLong(),
'                                       SetWindowLong(), SetLayeredWindowAttributes(), SetFormOpacity())
'               BLC, 3/17/2016 -1.10 - added SetControlBackcolor(), CTRL_DEFAULT_BACKCOLOR, Check1000hrFuels
'               BLC, 3/29/2016 -1.11 - added SetControlHighlight()
'               BLC, 4/1/2016 - 1.12 - added AddTallyValue()
'               BLC, 3/22/2017 - 1.13 - added SortListForm() from big rivers,
'                                       moved to mod_Forms (6/1/2016 big rivers dev):
'                                       CaptureEscapeKey(), SetFormOpacity()
'               BLC, 3/23/2017 - 1.14 - added PopulateForm(), DeleteRecord() from big rivers
'               BLC, 3/30/2017 - 1.15 - moved DeleteRecord() to mod_Db
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------
' -- Constants --
Private Const LWA_ALPHA     As Long = &H2
Private Const GWL_EXSTYLE   As Long = -20
Private Const WS_EX_LAYERED As Long = &H80000

Public Const CTRL_DEFAULT_BACKCOLOR  As Long = 65535  'RGB(255, 255, 0) highlight yellow

' -- Values --
Public NoData As Scripting.Dictionary

' -- Functions --
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long) As Long
 
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
  (ByVal hWnd As Long, _
   ByVal nIndex As Long, _
   ByVal dwNewLong As Long) As Long
 
Private Declare Function SetLayeredWindowAttributes Lib "user32" _
  (ByVal hWnd As Long, _
   ByVal crKey As Long, _
   ByVal bAlpha As Byte, _
   ByVal dwFlags As Long) As Long

' =================================
' SUB:          RollupReportbyPark
' Description:  Prepares concatenated report data
'               Looks for the number of records (years) for each ParkPlotSpecies (species found on a given park plot)
'               and concatenates the years (e.g. 2008|2009|2013 ) so that a species only takes up a single
'               row for a specific park plot in the report. This reduces report length by 50% or more.
' Assumptions:  Assumes that tlu_NCPN_Plants contains Utah_Species names for all species
'               identified in the plots. Also assumes temp_Sp_Rpt_by_Park_Complete has been run prior to
'               running this so the report is updated with the most recent data.
' Note:         -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, August 27, 2015 - for NCPN tools
' Revisions:    BLC, 8/27/2015 - initial version
' =================================
Public Sub RollupReportbyPark()
On Error GoTo Err_Handler

    Dim strParkPlotSpecies As String, strSpeciesYears As String
    Dim strPark As String, strFamily As String, strUtah_Species As String, strParkPlot As String
    Dim intPlotID As Integer, i As Integer, iCount As Integer
    Dim rs As DAO.Recordset, rsTemp As DAO.Recordset, rsCount As DAO.Recordset
    'Dim blnAdd As Boolean
    'Dim strSpeciesYr As String
    Dim strSQL As String
    
    Dim strPrevParkPlotSpecies As String
    
    
    Set rs = CurrentDb.OpenRecordset("temp_Sp_Rpt_by_Park_Complete")

    'remove existing table
    If DCount("[Name]", "MSysObjects", "[Name] = 'temp_Sp_Rpt_by_Park_Rollup'") = 1 Then _
            CurrentDb.TableDefs.Delete "temp_Sp_Rpt_by_Park_Rollup"
    
    'create empty table
    CreateRollupTable
    Set rsTemp = CurrentDb.OpenRecordset("temp_Sp_Rpt_by_Park_Rollup")
    
    'default
    strParkPlotSpecies = ""
    strSpeciesYears = ""
    'blnAdd = False
    
    If Not (rs.EOF And rs.BOF) Then
        rs.MoveFirst
        
        Do Until rs.EOF
            
            'set the current record's values
            strPark = rs("Unit_Code")
            intPlotID = rs("Plot_ID")
            strFamily = rs("Master_Family")
            strUtah_Species = rs("Utah_Species")
            strParkPlotSpecies = rs("ParkPlotSpecies")
            strParkPlot = rs("ParkPlot")
            'strSpeciesYr = rs("Year")
            
            If Not iCount > 0 Then
              'determine how many have the same ParkPlotSpecies
              strSQL = "SELECT COUNT(Year) AS NumRecords FROM temp_Sp_Rpt_by_Park_Complete WHERE ParkPlotSpecies = '" & strParkPlotSpecies & "';"
              Set rsCount = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
              iCount = rsCount!NumRecords
            End If
          
            For i = 1 To iCount
              'add year if it's a new year
              If Len(strSpeciesYears) = Len(Replace(strSpeciesYears, CStr(rs("Year")), "")) Then
                  strSpeciesYears = IIf(Len(strSpeciesYears) > 0, strSpeciesYears & "|" & rs("Year"), rs("Year"))
              End If
              rs.MoveNext
            Next
            
            ' add new record
            With rsTemp
                .AddNew
                !Unit_Code = strPark
                !Plot_ID = intPlotID
                !Master_Family = strFamily
                !Utah_Species = strUtah_Species
                !SpeciesYears = IIf(Len(strSpeciesYears) > 0, strSpeciesYears, rs!Year)
                !PlotParkSpecies = strParkPlotSpecies
                !ParkPlot = strParkPlot
                'update when rs!ParkPlotSpecies <> strParkPlotSpecies
                .Update
            End With
            'reset values
            strSpeciesYears = ""
            iCount = 0
        Loop
    End If
    
Exit_Sub:
    Set rs = Nothing
    Set rsTemp = Nothing
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - RollupReportbyPark[mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' =================================
' SUB:          CreateRollupTable
' Description:  Prepares rollup temporary table
' Assumptions:
' Note:         -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
' Source/date:  Bonnie Campbell, August 27, 2015 - for NCPN tools
' Revisions:    BLC, 8/27/2015 - initial version
' =================================
Public Sub CreateRollupTable()
On Error GoTo Err_Handler

    Dim tdf As DAO.TableDef
    
    Set tdf = CurrentDb.CreateTableDef("temp_Sp_Rpt_by_Park_Rollup")
    
    'add the new record
    With tdf
        .Fields.Append .CreateField("Unit_Code", dbText)
        .Fields.Append .CreateField("Plot_ID", dbInteger)
        .Fields.Append .CreateField("Master_Family", dbText)
        .Fields.Append .CreateField("Utah_Species", dbText)
        .Fields.Append .CreateField("SpeciesYears", dbText)
        .Fields.Append .CreateField("PlotParkSpecies", dbText)
        .Fields.Append .CreateField("ParkPlot", dbText)
    End With

    CurrentDb.TableDefs.Append tdf
    
Exit_Sub:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateRollupTable[mod_App_UI])"
    End Select
    Resume Exit_Sub
End Sub

' ---------------------------------
' SUB:          GetNoDataCollected
' Description:  Gets no data collected information from NoDataCollected table for event ID
' Assumptions:  -
' Parameters:   levelID - ID for event or event|transect as appropriate
'               level - event or transect (E = event, T = transect)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 2/11/2016 - added level to accommodate both event & transect level identifiers
'   BLC, 3/18/2016 - added 1000hr fuel A-D to handle no fuels reported in comments for transects
' ---------------------------------
Public Function GetNoDataCollected(levelID As String, level As String) As Scripting.Dictionary
On Error GoTo Err_Handler

    Dim strSQL As String, strItem As String
    Dim rs As DAO.Recordset
    
    Set NoData = New Scripting.Dictionary 'publicly set
    
    'prepare default dictionary
    With NoData
        .Add "1mBelt-Shrub", 0
        .Add "1mBelt-TreeSeedling", 0
'        .Add "1mBelt-ExoticPerennial", 0
        .Add "1mBelt-Exotics", 0
        .Add "OverstoryTree-Sapling", 0
        .Add "OverstoryTree-Census", 0
        .Add "Fuel-1000hr", 0
        .Add "Fuel-1000hr-A", 0
        .Add "Fuel-1000hr-B", 0
        .Add "Fuel-1000hr-C", 0
        .Add "Fuel-1000hr-D", 0
        .Add "SiteImpact-Disturbance", 0
        .Add "SiteImpact-Exotic", 0
    End With
    
    strSQL = "SELECT SampleType FROM NoDataCollected WHERE ID = '" & levelID & "' AND SampleLevel = '" & level & "';"
    
    Set rs = CurrentDb.OpenRecordset(strSQL)
    
    'rs.MoveFirst
    
    If Not (rs.EOF And rs.BOF) Then
    
        Do Until rs.EOF
    
            strItem = rs("SampleType") 'cannot use directly in NoData.item(rs("SampleType")) -> adds new item
            NoData.item(strItem) = 1
            
            rs.MoveNext
            
        Loop
        
    End If
    
    Set GetNoDataCollected = NoData
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - GetNoDataCollected[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     SetNoDataCollected
' Description:  Sets no data checkbox
' Assumptions:  Absolute value of Access/VBA checkbox is sent to drive 1 = true, 0 = false
'               SampleLevel is used vs. level in SQL (Access restricted word)
' Parameters:   levelID - ID for event/transect
'               level - sampling level identifier (E-event, T-transect)
'               SampleType - sub-protocol w/o data "1mBelt-Shrub", "OverstoryTree-Sapling", etc.
'               cbxValue - the value (1 or 0) to add or remove the
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:
' Adapted:      Bonnie Campbell, February 9, 2016 - for NCPN tools
' Revisions:
'   BLC, 2/9/2016  - initial version
'   BLC, 2/11/2016 - added level to accommodate both event & transect level identifiers
' ---------------------------------
Public Function SetNoDataCollected(levelID As String, level As String, SampleType As String, cbxValue As Integer) As Scripting.Dictionary
On Error GoTo Err_Handler
    
    Dim strSQL As String, strItem As String
    Dim rs As DAO.Recordset
    
    Set NoData = New Scripting.Dictionary 'publicly set
    Set NoData = GetNoDataCollected(levelID, level)
    
    NoData.item(SampleType) = cbxValue
    
    'update the table appropriately
    If cbxValue = 1 Then
        strSQL = "INSERT INTO NoDataCollected(ID, SampleLevel, SampleType) VALUES ('" & levelID & "', '" & level & "', '" & SampleType & "');"
    ElseIf cbxValue = 0 Then
        strSQL = "DELETE * FROM NoDataCollected WHERE ID = '" & levelID & "' AND SampleLevel = '" & level & _
                    "' AND SampleType = '" & SampleType & "';"
    End If
    
    DoCmd.SetWarnings (False)
    DoCmd.RunSQL (strSQL)
    DoCmd.SetWarnings (True)
    
    'return current dictionary object
    Set SetNoDataCollected = NoData
    
Exit_Handler:
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetNoDataCollected[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' SUB:          SetControlBackcolor
' Description:  sets controls backcolor based on control value
' Parameters:   ctrl - textbox control (textbox)
'               threshold - value to compare against (variant)
'               compareType - type of comparison (string)
'               color - numeric value for color (long) - result of RGB(r,g,b)
'               checkNULL - check if the control's value is NULL (boolean)
'               checkEmpty - check if the control's value is an empty string (boolean)
' Returns:      -
' Assumptions:  Assumes CTRL_DEFAULT_BACKCOLOR is set for the application
'               and that this is the typical backcolor for the controls
'               using SetControlBackcolor.
'               Assumes threshold value is numeric.
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/17/2016 - initial version
' ---------------------------------
Public Sub SetControlBackcolor(ctrl As TextBox, color As Long, checkNULL As Boolean, _
                        checkEmpty As Boolean, Optional threshold As Variant, Optional compareType As String)
On Error GoTo Err_Handler
    
    Dim resetcolor As Boolean
    
    'default
    resetcolor = False
    
    'change the backcolor --> revert to default only if the conditions aren't met
    ctrl.BackColor = color
    
    'null
    If checkNULL Then
        'reset backcolor if null
        If IsNull(Trim(ctrl.Text)) Then
            resetcolor = True
            GoTo Exit_Handler
        End If
    End If
    
    'empty
    If checkEmpty Then
        'reset backcolor if empty
        If Len(Trim(ctrl.Text)) = 0 Then
            resetcolor = True
            GoTo Exit_Handler
        End If
    End If
    
    'threshold
    If Not IsNull(threshold) And IsNumeric(ctrl.Text) Then
        'set value base on compareType & threshold
        Select Case compareType
            Case "gt"
                If Not CDbl(ctrl.Text) > threshold Then
                    resetcolor = True
                End If
            Case "gteq"
                If Not CDbl(ctrl.Text) >= threshold Then
                    resetcolor = True
                End If
            Case "lt"
                If Not CDbl(ctrl.Text) < threshold Then
                    resetcolor = True
                End If
            Case "lteq"
                If Not CDbl(ctrl.Text) <= threshold Then
                    resetcolor = True
                End If
            Case "eq"
                If Not CDbl(ctrl.Text) = threshold Then
                    resetcolor = True
                End If
        End Select
    End If
    
Exit_Handler:
    'reset to default backcolor
    If resetcolor Then
        ctrl.BackColor = CTRL_DEFAULT_BACKCOLOR
    End If
    
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetControlBackcolor[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          Check1000hrFuels
' Description:  Handles 1000hr fuel check actions
' Assumptions:  -
' Parameters:   -
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, March 18, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 3/18/2016  - initial version
'   BLC, 3/23/2016  - remove setting values when no records found
' ---------------------------------
Public Sub Check1000hrFuels()
On Error GoTo Err_Handler

    Dim frm As Form
    Set frm = Forms!frm_Data_Entry!fsub_Fuels_1000.Form

    '-----------------------------------
    ' update the NoDataCollected info IF no records now exist
    '-----------------------------------
    If frm.RecordsetClone.RecordCount = 0 Then
    
'        Dim NoData As Scripting.Dictionary
'
'        With frm.Parent.Form
'            'add the no data collected record
'            Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr", 1)
'
'            'update checkbox/rectangle --> No1000hr is not set here (leave commented out)
''            .Controls("cbxNo1000hr") = 1
''            .Controls("cbxNo1000hr").Enabled = True
''            .Controls("rctNo1000hr").Visible = True
'
'            'update A, B, C, D transect 1000hr fuels as well
'            .Controls("cbxNo1000hrA") = 1
'            .Controls("cbxNo1000hrA").Enabled = True
'            .Controls("rctNo1000hrA").Visible = True
'
'            .Controls("cbxNo1000hrB") = 1
'            .Controls("cbxNo1000hrB").Enabled = True
'            .Controls("rctNo1000hrB").Visible = True
'
'            .Controls("cbxNo1000hrC") = 1
'            .Controls("cbxNo1000hrC").Enabled = True
'            .Controls("rctNo1000hrC").Visible = True
'
'            .Controls("cbxNo1000hrD") = 1
'            .Controls("cbxNo1000hrD").Enabled = True
'            .Controls("rctNo1000hrD").Visible = True
'
'            'add the database records for A-D
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-A", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-B", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-C", 1
'            SetNoDataCollected .Controls("Event_ID"), "E", "Fuel-1000hr-D", 1
'        End With
        
    Else
    
        'default values
        With frm.Parent.Form
            'update checkbox/rectangle (leave 1000hr commented here)
'            .Controls("cbxNo1000hr") = 0
'            .Controls("cbxNo1000hr").Enabled = True
'            .Controls("rctNo1000hr").Visible = True
            
            'update A, B, C, D transect 1000hr fuels as well
            .Controls("cbxNo1000hrA") = 0
            .Controls("cbxNo1000hrA").Enabled = True
            .Controls("rctNo1000hrA").Visible = True
            
            .Controls("cbxNo1000hrB") = 0
            .Controls("cbxNo1000hrB").Enabled = True
            .Controls("rctNo1000hrB").Visible = True
            
            .Controls("cbxNo1000hrC") = 0
            .Controls("cbxNo1000hrC").Enabled = True
            .Controls("rctNo1000hrC").Visible = True
        
            .Controls("cbxNo1000hrD") = 0
            .Controls("cbxNo1000hrD").Enabled = True
            .Controls("rctNo1000hrD").Visible = True
        End With
    
        'check for A, B, C, D transect 1000hr fuels
        Dim rs As DAO.Recordset
        
        Set rs = frm.RecordsetClone
        With rs
            .MoveFirst
            Do While Not .EOF
            Select Case .Fields("Transect")
            
                Case "A"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-A", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrA") = 0
                        .Controls("cbxNo1000hrA").Enabled = False
                        .Controls("rctNo1000hrA").Visible = False
                    End With
                    
                Case "B"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-B", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrB") = 0
                        .Controls("cbxNo1000hrB").Enabled = False
                        .Controls("rctNo1000hrB").Visible = False
                    End With
                    
                Case "C"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-C", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrC") = 0
                        .Controls("cbxNo1000hrC").Enabled = False
                        .Controls("rctNo1000hrC").Visible = False
                    End With
                    
                Case "D"
                    With frm.Parent.Form
                        'remove the no data collected record
                        Set NoData = SetNoDataCollected(.Controls("Event_ID"), "E", "Fuel-1000hr-D", 0)
                            
                        'update checkbox/rectangle
                        .Controls("cbxNo1000hrD") = 0
                        .Controls("cbxNo1000hrD").Enabled = False
                        .Controls("rctNo1000hrD").Visible = False
                    End With
            End Select
            .MoveNext
            Loop
        End With
        
        'set checkboxes based on NoDataCollected (catch unchanged checkboxes)
        Dim dNoDataEvent As Scripting.Dictionary
        Set dNoDataEvent = GetNoDataCollected(frm.Parent.Form.Controls("Event_ID"), "E")
        
        With dNoDataEvent
            frm.Parent.Form.Controls("cbxNo1000hrA") = .item("Fuel-1000hr-A")
            frm.Parent.Form.Controls("cbxNo1000hrB") = .item("Fuel-1000hr-B")
            frm.Parent.Form.Controls("cbxNo1000hrC") = .item("Fuel-1000hr-C")
            frm.Parent.Form.Controls("cbxNo1000hrD") = .item("Fuel-1000hr-D")
        End With
    End If

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Check1000hrFuels[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          SetControlHighlight
' Description:  handles control highlight actions
' Parameters:   ctrl - textbox control (textbox)
'               -- optional --
'               threshold - value to compare control value to (double, default = 0)
'               compareType - how control value should be compared to threshold (string, default = "gteq")
' Returns:      -
' Assumptions:  highlighting will be consistent across all textboxes
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, March 2016
' Revisions:    BLC, 3/29/2016 - initial version
' ---------------------------------
Public Sub SetControlHighlight(ctrl As TextBox, Optional threshold As Double, Optional compareType As String)
On Error GoTo Err_Handler

    'set defaults if optional values aren't set
    If Not IsNumeric(threshold) Then threshold = 0
    If Len(compareType) > 0 Then compareType = "gteq"

    'set the backcolor to white when the value reaches a threshold >= 0, checking for NULL and empty values
    SetControlBackcolor ctrl, RGB(255, 255, 255), True, True, threshold, compareType
   
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SetControlHighlight[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          AddTallyValue
' Description:  Adds tally amount to control
' Assumptions:  -
' Parameters:   ctrl - control being changed (textbox)
'               tallyAmount - amount to add (integer - positive or negative)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Public Sub AddTallyValue(ctrl As TextBox, tallyAmount As Integer)
On Error GoTo Err_Handler
  
  'handle when the user keeps cursor in field & tallyAmount would drive the value to < 0 (negative)
  If (ctrl.Value + tallyAmount < 0) Or (IsNull(ctrl.Value) And tallyAmount < 0) Then GoTo Exit_Handler
  
  If tallyAmount = 0 Then ctrl.Value = 0
  
  Select Case ctrl.Name
    Case "SeedTotal"
        ctrl.Value = Nz(ctrl.Value, 0) + tallyAmount
  End Select
  
  'return focus
  ctrl.SetFocus
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - AddTallyValue[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' SUB:          DisableTallyButtons
' Description:  Disable tally buttons on control
' Assumptions:  -
' Parameters:   frm - form where tally buttons are being changed (form)
'               lookFor - common part of tally button name (string)
' Returns:      N/A
' Throws:       none
' References:   none
' Source/date:  Bonnie Campbell, April 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 4/1/2016  - initial version
' ---------------------------------
Public Sub DisableTallyButtons(frm As Form, lookFor As String)
On Error GoTo Err_Handler
  
  Dim ctrl As Control
  
  For Each ctrl In frm.Controls
  
    If Len(ctrl.Name) > Len(Replace(ctrl.Name, lookFor, "")) Then
    
            ctrl.Enabled = False
    
    End If
  
  Next
  
Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - DisableTallyButtons[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          SortListForm
' Description:  form label sort on click actions
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:
'   pere_de_chipstic, August 5, 2012
'   http://www.utteraccess.com/forum/Sort-Continuous-Form-Hea-t1991553.html
'   Allen Browne, June 28, 2006
'   https://bytes.com/topic/access/answers/506322-using-orderby-multiple-fields
' Source/date:  Bonnie Campbell, January 19, 2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 1/19/2017 - initial version
'   BLC - 1/31/2017 - adjusted to accommodate templates list
'   BLC - 2/21/2017 - adjusted to accommodate Contact list
' ---------------------------------
Public Sub SortListForm(frm As Form, ctrl As Control)
On Error GoTo Err_Handler

    Dim strSort As String
    
    'default
    strSort = ""
    
    'set sort field
    Select Case Replace(ctrl.Name, "lbl", "")
        Case "Email"
            strSort = "Email"
        Case "HdrID"
            strSort = "ID"
            Select Case frm.Name
                Case "ContactList"
                    strSort = "c.ID"
            End Select
        Case "Name"
            strSort = "LastName"
        Case "Template"
            strSort = "TemplateName"
        Case "SOPNum"
            strSort = "SOPNumber"
        Case "SOP"
            strSort = "FullName"
        Case "Syntax"
            strSort = "Syntax"
        Case "Version"
            strSort = "Version"
        Case "EffectiveDate"
            strSort = "EffectiveDate"
        Case ""
    End Select

    'set the sort
    If InStr(frm.OrderBy, strSort) = 0 Then
        frm.OrderBy = strSort
    ElseIf Right(frm.OrderBy, 4) = "Desc" Then
        frm.OrderBy = strSort
    Else
        frm.OrderBy = strSort & " Desc"
    End If
    
    frm.OrderByOn = True
    
Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - SortListForm[mod_App_UI form])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          PopulateForm
' Description:  Populate a form using a specific record for edits
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, June 1, 2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 6/1/2016 - initial version
'   BLC - 6/2/2016 - moved from forms (EventsList, TaglineList)
'   BLC - 8/8/2016 - revised to use default table name
'   BLC - 8/29/2016 - adjusted for Contact form (requires both Contact, Contact_Access data)
'                     using usys_temp_qdf & adjusting ID to Contact_ID in final SQL
'   BLC - 10/24/2016 - added ModWentworth form
'   BLC - 1/12/2017 - code cleanup
'   BLC - 2/14/2017 - added Task form
' --------------------------------------------------------------------
'   BLC - 3/23/2017 - adapted version for Upland db
' --------------------------------------------------------------------
' ---------------------------------
Public Sub PopulateForm(frm As Form, ID As Long)
On Error GoTo Err_Handler
    Dim strSQL As String, strTable As String

    With frm
        'default
        strTable = .Name
        
        'find the form & populate its controls from the ID
        Select Case .Name
            Case "Contact"
                'requires Contact & Contact_Access data
                Dim qdf As DAO.QueryDef
                CurrentDb.QueryDefs("usys_temp_qdf").sql = GetTemplate("s_contact_access")
                
                strTable = "usys_temp_qdf"
                'set form fields to record fields as datasource
                'contact data
                .Controls("tbxID").ControlSource = "c.ID"
                .Controls("tbxFirst").ControlSource = "FirstName"
                .Controls("tbxMI").ControlSource = "MiddleInitial"
                .Controls("tbxLast").ControlSource = "LastName"
                .Controls("tbxEmail").ControlSource = "Email"
                .Controls("tbxUsername").ControlSource = "Username"
                .Controls("tbxOrganization").ControlSource = "Organization"
                .Controls("tbxPhone").ControlSource = "WorkPhone"
                .Controls("tbxPosition").ControlSource = "PositionTitle"
                .Controls("tbxExtension").ControlSource = "WorkExtension"
                'contact_access data
                .Controls("cbxUserRole").ControlSource = "Access_ID"
            Case "Events"
                strTable = "Event"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("cbxSite").ControlSource = "Site_ID"
                .Controls("cbxLocation").ControlSource = "Location_ID"
                .Controls("tbxStartDate").ControlSource = "StartDate"
                .Controls("lblMsgIcon").Caption = ""
                .Controls("lblMsg").Caption = ""
            Case "Feature"
                'set form fields to record fields as datasource
                .Controls("tbxID").ControlSource = "ID"
                .Controls("tbxFeature").ControlSource = "Feature"
                '.Controls("cbxLocation").ControlSource = ""
        End Select
    
'        'save record changes from form first to avoid "Write Conflict" errors
'        'where form & SQL are attempting to save record
'        'frm.Dirty = False
'
'        If frm.Dirty Then
'            MsgBox frm.Name & " DIRTY"
'            frm.Dirty = False
'        Else
'            MsgBox frm.Name & " CLEAN"
'        End If
        
        strSQL = GetTemplate("s_form_edit", "tbl" & PARAM_SEPARATOR & strTable & "|id" & PARAM_SEPARATOR & ID)
        
        'alter to retrieve proper ID
        Select Case .Name
            Case "Contact"
                strSQL = Replace(strSQL, " ID = ", " c.ID = ")
        End Select
        
        .RecordSource = strSQL
        
    End With

Exit_Handler:
    Exit Sub
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - PopulateForm[mod_App_UI])"
    End Select
    Resume Exit_Handler
End Sub