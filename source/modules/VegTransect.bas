Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegTransect
' Level:        Framework class
' Version:      1.09
'
' Description:  VegTransect object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 10/28/2015 - 1.00 - initial version
'               BLC - 8/8/2016   - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               BLC - 7/5/2017   - 1.02 - AddQuadrats() & AddSurfaceMicrohabitats() to
'                                         initialize records tied to new transect
'               BLC - 7/13/2017  - 1.03 - added UpdateTransect for updating visit data,
'                                         added StartTime, Comments properties
'               BLC - 7/16/2017  - 1.04 - revised to accommodate NULL StartTime values
'               BLC - 7/25/2017  - 1.05 - revised GetTransectQuadrats() to address empty recordsets for transects w/o quadrats
' --------------------------------------------------------------------------------------
'               BLC - 8/23/2017  - 1.06 - merge in prior work:
'                                                                                       Big Rivers park casing for Transect property
' --------------------------------------------------------------------------------------
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.07 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/6/2017  - 1.08 - removed GetClass() after Factory class instatiation implemented
'               BLC - 11/6/2017  - 1.09 - add site_vegtransect linkage (SaveToDb())
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_LocationID As Long
Private m_EventID As Long
Private m_TransectQuadratID As String
Private m_TransectNumber As Integer
Private m_SampleDate As Date

Private m_StartTime As Date
Private m_Comments As String

Private m_Park As String
Private m_ObserverID As Integer
Private m_RecorderID As Integer
Private m_ObserverName As String
Private m_Observer As String
Private m_RecorderName As String
Private m_Recorder As String

Private m_HasQuadrats As Boolean
Private m_TransectQuadrats As Variant
Private m_NumQuadrats As Integer

'---------------------
' Events
'---------------------
Public Event InvalidTransectNumber(Value As Integer)
Public Event InvalidTransectQuadratID(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
    'set the appropriate park value
'    Me.Park = GetParkCode(Value)
End Property

Public Property Get LocationID() As Long
    LocationID = m_LocationID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let TransectQuadratID(Value As String)
    m_TransectQuadratID = Value
    
    'set the tempvar also
    SetTempVar "TransectQuadratID", Value
    
    'populate related properties
    GetTransectQuadrats
End Property

Public Property Get TransectQuadratID() As String
    TransectQuadratID = m_TransectQuadratID
End Property

Public Property Let TransectNumber(Value As Integer)
    If IsNull(Me.Park) Then
        MsgBox "Park must be set before setting transect number.", vbCritical, "Missing Park"
        
    End If
    'validate park (BLCA & CANY only)
'    Select Case Me.Park
'        Case "BLCA", "CANY"
'            'check value
'            'validate transect #
'            Dim aryTransectNum() As String
'            aryTransectNum = Split(TRANSECT_NUMBERS, ",")
'            If IsInArray(CStr(value), aryTransectNum) Then
'                m_TransectNumber = value
'            Else
'                RaiseEvent InvalidTransectNumber(value)
'            End If
'        Case "DINO"
'            'invalid
'            RaiseEvent InvalidTransectNumber(value)
'        Case Else
'            'invalid
'            RaiseEvent InvalidTransectNumber(value)
'    End Select
End Property

Public Property Get TransectNumber() As Integer
    TransectNumber = m_TransectNumber
End Property

Public Property Let SampleDate(Value As Date)
    m_SampleDate = Value
End Property

Public Property Get SampleDate() As Date
    SampleDate = m_SampleDate
End Property

Public Property Let Park(Value As String)
    m_Park = Value
End Property

Public Property Get Park() As String
    Park = m_Park
End Property

Public Property Let ObserverID(Value As Integer)
    m_ObserverID = Value
End Property

Public Property Get ObserverID() As Integer
    ObserverID = m_ObserverID
End Property

Public Property Let Observer(Value As String)
    m_Observer = Value
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let ObserverName(Value As String)
    m_ObserverName = Value
End Property

Public Property Get ObserverName() As String
    ObserverName = m_ObserverName
End Property

Public Property Let RecorderID(Value As Integer)
    m_RecorderID = Value
End Property

Public Property Get RecorderID() As Integer
    RecorderID = m_RecorderID
End Property

Public Property Let Recorder(Value As String)
    m_Recorder = Value
End Property

Public Property Get Recorder() As String
    Recorder = m_Recorder
End Property

Public Property Let RecorderName(Value As String)
    m_RecorderName = Value
End Property

Public Property Get RecorderName() As String
    RecorderName = m_RecorderName
End Property

Public Property Let HasQuadrats(Value As Boolean)
    m_HasQuadrats = Value
End Property

Public Property Get HasQuadrats() As Boolean
    HasQuadrats = m_HasQuadrats
End Property

Public Property Let TransectQuadrats(Value As Variant)
    m_TransectQuadrats = Value
End Property

Public Property Get TransectQuadrats() As Variant
    TransectQuadrats = m_TransectQuadrats
End Property

Public Property Let NumQuadrats(Value As Integer)
    m_NumQuadrats = Value
End Property

Public Property Get NumQuadrats() As Integer
    NumQuadrats = m_NumQuadrats
End Property

Public Property Let StartTime(Value As Date)
    m_StartTime = Value
End Property

Public Property Get StartTime() As Date
    StartTime = m_StartTime
End Property

Public Property Let Comments(Value As String)
    m_Comments = Value
End Property

Public Property Get Comments() As String
    Comments = m_Comments
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========
' handled by Factory class

'======== Standard Methods ==========

' ---------------------------------
' SUB:          Class_Initialize
' Description:  Initialize the class
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  -
' Adapted:      Bonnie Campbell, April 4, 2016 - for NCPN tools
' Revisions:
'   BLC - 4/4/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          Class_Terminate
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

    'Set m_ID = 0

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 4/4/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 9/8/2016 - code cleanup
'   BLC, 11/6/2017 - add site_vegtransect linkage
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegtransect"
    
    Dim Params(0 To 6) As Variant

    With Me
        Params(0) = "VegTransect"
        Params(1) = .LocationID
        Params(2) = .EventID
        Params(3) = .TransectNumber
        Params(4) = .SampleDate
        
        If IsUpdate Then
            Template = "u_vegtransect"
            Params(5) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
    'SetObserverRecorder Me, "VegTransect"
    
    'add big rivers transect Site_VegTransect linking record
    If APP = "Big_Rivers" Then
    
        Params(0) = "Site_VegTransect"
        Params(1) = GetSiteID(TempVars("ParkCode"), TempVars("SiteCode"))   'Site ID
        Params(2) = Me.ID   'VegTransect ID
        SetRecord "i_site_vegtransect", Params
    
    End If


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' Function:     GetTransectQuadrats
' Description:  Fetch the quadrats for the transect (if any)
' Parameters:   -
' Returns:      Array of quadrats
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/10/2017 - for NCPN tools
' Revisions:
'   BLC, 7/10/2017 - initial version
'   BLC, 7/25/2017 - revised to address empty recordsets for transects w/o quadrats
'---------------------------------------------------------------------------------------
Public Function GetTransectQuadrats() As Variant
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    
    'retrieve array of all quadrats associated w/ this transect
    Set rs = GetRecords("s_transect_quadrat_IDs")
    
    With Me
        'defaults
        .HasQuadrats = False
        .NumQuadrats = 0
        
        'ensure there are records before moves
        If Not (rs.BOF And rs.EOF) Then
            rs.MoveLast
            rs.MoveFirst
            
            'set had quadrats
            .HasQuadrats = True
            .NumQuadrats = rs.RecordCount
            
            'return the 2-dimensional array (1-columns, 2-rows)
            .TransectQuadrats = rs.GetRows(rs.RecordCount)
            
        Else
        
            'return 0 (no records)
            .TransectQuadrats = rs.RecordCount
        
        End If
            
        GetTransectQuadrats = .TransectQuadrats
    
    End With
    
Exit_Handler:
    rs.Close
    Set rs = Nothing
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetTransectQuadrats[VegTransect class])"
    End Select
    Resume Exit_Handler
End Function

'---------------------------------------------------------------------------------------
' SUB:          AddQuadrats
' Description:  Adds quadrats to a new transect which has no quadrat records
'               QUADRATS_PER_TRANSECT (see mod_App_Settings) is the current # of quadrats
'               existing along a transect

' Parameters:   QuadratNum - number of quadrat to add (integer, optional)
'                   0 - add all quadrats 1 - QUADRATS_PER_TRANSECT value (currently 3)
'                   1, 2, or 3 - add quadrat 1,2, or 3
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/5/2017 - for NCPN tools
' Revisions:
'   BLC, 7/5/2017 - initial version
'   BLC, 7/18/2017 - replace 3 with QUADRATS_PER_TRANSECT
'---------------------------------------------------------------------------------------
Public Sub AddQuadrats(Optional QuadratNum As Integer = 0)
On Error GoTo Err_Handler
    
    Dim Template As String
    Dim i As Integer
    
    Template = "i_new_transect_quadrat"
    
    Dim Params(0 To QUADRATS_PER_TRANSECT) As Variant

    'if QuadratNum is set
    If QuadratNum <> 0 Then
        With Me
            Params(0) = "Transect"
            Params(1) = .TransectQuadratID
            Params(2) = i                   'quadrat number
            
            .ID = SetRecord(Template, Params)
        End With
    
        'exit
        GoTo Exit_Handler
    End If

    'if QuadratNum = 0 then assume add all quadrats
    'iterate once per quadrat
    For i = 1 To QUADRATS_PER_TRANSECT

        With Me
            Params(0) = "Transect"
            Params(1) = .TransectQuadratID
            Params(2) = i                   'quadrat number
            
            .ID = SetRecord(Template, Params)
        End With
        
        'SetObserverRecorder Me, "VegTransect"
    
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - AddQuadrats[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          AddSurfaceMicrohabitats
' Description:  Adds quadrats to a new transect which has no quadrat records
'               QUADRATS_PER_TRANSECT (see mod_App_Settings) is the current # of quadrats
'               existing along a transect

' Parameters:   QuadratNum - number of quadrat to add (integer, optional)
'                   0 - add all quadrats 1 - QUADRATS_PER_TRANSECT value (currently 3)
'                   1, 2, or 3 - add quadrat 1,2, or 3
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/5/2017 - for NCPN tools
' Revisions:
'   BLC, 7/5/2017 - initial version
'   BLC, 7/18/2017 - replace 3 with QUADRATS_PER_TRANSECT
'---------------------------------------------------------------------------------------
Public Sub AddSurfaceMicrohabitats(Optional SfcMicrohabitat As Integer = 0)
On Error GoTo Err_Handler
    
    Dim Template As String
    Dim arySurfaces As Variant
    Dim aryQuadrats As Variant
    Dim rs As DAO.Recordset
    Dim sfc_id As Variant
    Dim QuadratID As Variant
    
    'retrieve array of all surface IDs
    Set rs = GetRecords("s_surface_IDs")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        rs.MoveFirst
        arySurfaces = rs.GetRows(rs.RecordCount)
    Else
        arySurfaces = 0
    End If
    
    'retrieve array of all quadrats associated w/ this transect
    Set rs = GetRecords("s_transect_quadrat_IDs")
    
    If Not (rs.BOF And rs.EOF) Then
        rs.MoveLast
        rs.MoveFirst
        aryQuadrats = rs.GetRows(rs.RecordCount)
    Else
        aryQuadrats = 0
    End If
    
    Template = "i_new_transect_quadrat_sfccover"
    
    Dim Params(0 To QUADRATS_PER_TRANSECT) As Variant

    'iterate once per quadrat
    For Each QuadratID In aryQuadrats
    
        'iterate once per surface
        For Each sfc_id In arySurfaces
    
            With Me
                Params(0) = "Transect"
                Params(1) = QuadratID       'quadrat ID
                Params(2) = sfc_id          'surface microhabitat ID
                
                .ID = SetRecord(Template, Params)
            End With
        
            'SetObserverRecorder Me, "VegTransect"
                
        Next
        
    Next
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - AddSurfaceMicrohabitats[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateObserver
' Description:  Updates transect observer
' Assumption:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/17/2017 - for NCPN tools
' Revisions:
'   BLC, 7/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub UpdateObserver()
On Error GoTo Err_Handler
    
    Dim Template As String
        
    Template = "u_transect_observer"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "Transect"
        Params(1) = .Observer           'observer
        Params(2) = .TransectQuadratID  'string identifier
        
        .ID = SetRecord(Template, Params)
                
    End With

    'SetObserverRecorder Me, "VegTransect"
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateObserver[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateStartTime
' Description:  Updates visit information (start time) for a transect
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/16/2017 - for NCPN tools
' Revisions:
'   BLC, 7/16/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub UpdateStartTime()
On Error GoTo Err_Handler
    
    Dim Template As String
        
    Template = "u_transect_start_time"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "Transect"
        Params(1) = .StartTime          'start time
        Params(2) = .TransectQuadratID  'string identifier
        
        .ID = SetRecord(Template, Params)
    End With

    'SetObserverRecorder Me, "VegTransect"
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateStartTime[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateComments
' Description:  Updates transect comments
' Assumptions:
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/17/2017 - for NCPN tools
' Revisions:
'   BLC, 7/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub UpdateComments()
On Error GoTo Err_Handler
    
    Dim Template As String
        
    Template = "u_transect_comments"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "Transect"
        Params(1) = .Comments           'comments
        Params(2) = .TransectQuadratID  'string identifier
        
        .ID = SetRecord(Template, Params)
                
    End With

    'SetObserverRecorder Me, "VegTransect"
    
Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateComments[VegTransect class])"
    End Select
    Resume Exit_Handler
End Sub