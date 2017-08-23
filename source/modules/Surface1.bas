Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        Surface
' Level:        Framework class
' Version:      1.03
'
' Description:  Surface (microhabitat) object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/17/2017
' References:   -
' Revisions:    BLC - 4/17/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_SurfaceID As Long
Private m_SfcName As String
Private m_SfcDescription As String
Private m_OrigColumnName As String

'---------------------
' Events
'---------------------
Public Event InvalidID(Value As Long)
Public Event InvalidSfcID(Value As Long)
Public Event InvalidSfcName(Value As String)
Public Event InvalidSfcDescription(Value As String)
Public Event InvalidOrigColumnName(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    If varType(Value) = vbLong Then
        m_ID = Value
        'also set surfaceID value
        m_SurfaceID = Value
    Else
        RaiseEvent InvalidID(Value)
    End If
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let SurfaceID(Value As Long)
    If varType(Value) = vbLong Then
        m_SurfaceID = Value
    Else
        RaiseEvent InvalidSfcID(Value)
    End If
End Property

Public Property Get SurfaceID() As Long
    SurfaceID = m_SurfaceID
End Property

Public Property Let SfcName(Value As String)
    'valid length varchar(25) or ZLS
    If IsBetween(Len(Value), 1, 25, True) Then
        m_SfcName = Value
    Else
        RaiseEvent InvalidSfcName(Value)
    End If
End Property

Public Property Get SfcName() As String
    SfcName = m_SfcName
End Property

Public Property Let SfcDescription(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_SfcDescription = Value
    Else
        RaiseEvent InvalidSfcDescription(Value)
    End If
End Property

Public Property Get SfcDescription() As String
    SfcDescription = m_SfcDescription
End Property

Public Property Let OrigColumnName(Value As String)
    'valid length varchar(25) or ZLS
    If IsBetween(Len(Value), 1, 25, True) Then
        m_OrigColumnName = Value
    Else
        RaiseEvent InvalidOrigColumnName(Value)
    End If
End Property

Public Property Get OrigColumnName() As String
    OrigColumnName = m_OrigColumnName
End Property

'---------------------
' Methods
'---------------------

' ---------------------------------
' Sub:          Class_Initialize
' Description:  Class initialization (starting) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[Surface class])"
    End Select
    Resume Exit_Handler
End Sub

' ---------------------------------
' Sub:          Class_Terminate
' Description:  Class termination (closing) event
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   -
' Source/date:  Bonnie Campbell, October 30, 2015 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC - 10/30/2015 - initial version
' ---------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_WoodyCanopy])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup surface based on surface/microhabitat ID
' Parameters:   ID - identifier for surface/microhabitat record (long)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/17/2017 - for NCPN tools
' Revisions:
'   BLC, 4/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(ID As Long)
On Error GoTo Err_Handler
    
    Dim rs As DAO.Recordset
    
    'set ID for parameters
    SetTempVar "SurfaceID", ID
    
    Set rs = GetRecords("s_surface_by_ID")
    If Not (rs.EOF And rs.BOF) Then
        With rs
            Me.ID = Nz(.Fields("ID"), 0)
            Me.SfcName = Nz(.Fields("Surface"), "")
            Me.SfcDescription = Nz(.Fields("Description"), "")
            Me.OrigColumnName = Nz(.Fields("ColName"), "")
        End With
    Else
        RaiseEvent InvalidID(ID)
    End If

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[Surface class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save surface/microhabitat based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/17/2017 - for NCPN tools
' Revisions:
'   BLC, 4/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_surface"
    
    Dim Params(0 To 5) As Variant

    With Me
        Params(0) = "Surface"
        Params(1) = .SfcName
        Params(2) = .SfcDescription
        Params(3) = .OrigColumnName
        
        If IsUpdate Then
            Template = "u_surface"
            Params(4) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[Surface class])"
    End Select
    Resume Exit_Handler
End Sub