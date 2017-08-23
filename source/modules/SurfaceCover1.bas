Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

' =================================
' CLASS:        SurfaceCover
' Level:        Framework class
' Version:      1.01
'
' Description:  Surface (microhabitat) cover object related properties, events, functions & procedures for UI display
'
' Source/date:  Bonnie Campbell, 4/17/2017
' References:   -
' Revisions:    BLC - 4/17/2017 - 1.00 - initial version
'               BLC - 4/24/2017 - 1.01 - revise PercentCover to Single vs. Integer
' =================================

'---------------------
' Declarations
'---------------------
Private m_Surface As New Surface

Private m_ID As Long

Private m_QuadratID As Long
Private m_SurfaceID As Long
Private m_PercentCover As Single

'---------------------
' Events
'---------------------
Public Event InvalidQuadratID(Value As Long)
Public Event InvalidPercentCover(Value As Single)

'-- base events (surface)
Public Event InvalidID(Value As Long)
Public Event InvalidSfcID(Value As Long)
Public Event InvalidSfcName(Value As String)
Public Event InvalidSfcDescription(Value As String)
Public Event InvalidOrigColumnName(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let QuadratID(Value As Long)
    m_QuadratID = Value
End Property

Public Property Get QuadratID() As Long
    QuadratID = m_QuadratID
End Property

Public Property Let PercentCover(Value As Single)
    If IsBetween(Value, 0, 100, True) Then
        m_PercentCover = Value
    Else
        RaiseEvent InvalidPercentCover(Value)
    End If
End Property

Public Property Get PercentCover() As Single
    PercentCover = m_PercentCover
End Property

' ---------------------------
' -- base class properties --
' ---------------------------
' NOTE: required since VBA does not support direct inheritance
'       or polymorphism like other OOP languages
' ---------------------------
' base class = Surface
' ---------------------------
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
        m_Surface.SfcName = Value
    Else
        RaiseEvent InvalidSfcName(Value)
    End If
End Property

Public Property Get SfcName() As String
    SfcName = m_Surface.SfcName
End Property

Public Property Let SfcDescription(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_Surface.SfcDescription = Value
    Else
        RaiseEvent InvalidSfcDescription(Value)
    End If
End Property

Public Property Get SfcDescription() As String
    SfcDescription = m_Surface.SfcDescription
End Property

Public Property Let OrigColumnName(Value As String)
    'valid length varchar(25) or ZLS
    If IsBetween(Len(Value), 1, 25, True) Then
        m_Surface.OrigColumnName = Value
    Else
        RaiseEvent InvalidOrigColumnName(Value)
    End If
End Property

Public Property Get OrigColumnName() As String
    OrigColumnName = m_Surface.OrigColumnName
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

    'MsgBox "Initializing...", vbOKOnly
    
    Set m_Surface = New Surface

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[cls_SurfaceCover])"
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
    
'    MsgBox "Terminating...", vbOKOnly
        
    Set m_Surface = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[cls_SurfaceCover])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup surface/microhabitat based on the identifier
' Parameters:   ID - identifier for surface/microhabitat (long
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

    m_Surface.Init (ID)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[cls_SurfaceCover])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save veg walk species based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/17/2017 - for NCPN tools
' Revisions:
'   BLC, 4/17/2017 - initial version, based on Big Rivers classes SaveToDb()
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
        
    Dim Template As String
    
    Template = "i_surface_cover"
    
    Dim Params(0 To 5) As Variant
    
    With Me
        Params(0) = "SurfaceCover"
        Params(1) = .QuadratID
        Params(2) = .SurfaceID
        Params(3) = .PercentCover
        
        If IsUpdate Then
            Template = "u_surface_cover"
            'Params(4) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
    'no RecordAction for invasives --> if added later see Big Rivers

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[cls_SurfaceCover])"
    End Select
    Resume Exit_Handler
End Sub