Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        SamplingEvent
' Level:        Framework class
' Version:      1.02
'
' Description:  Sampling event object related properties, events, functions & procedures for UI display
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 7/17/2017
' References:   -
' Revisions:    BLC - 7/17/2017 - 1.00 - initial version
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.01 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/6/2017  - 1.02 - removed GetClass() after Factory class instatiation implemented
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As String
Private m_StartDate As Date
Private m_Observer As String
Private m_Comments As String

'---------------------
' Events
'---------------------
Public Event InvalidID(Value As Long)
Public Event InvalidEventID(Value As String)
Public Event InvalidStartDate(Value As Date)
Public Event InvalidObserver(Value As String)
Public Event InvalidComments(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    If varType(Value) = vbLong Then
        m_ID = Value
    Else
        RaiseEvent InvalidID(Value)
    End If
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(Value As String)
    If varType(Value) = vbString Then
        m_EventID = Value
    
        'set ID for parameters
        SetTempVar "EventID", m_EventID
    Else
        RaiseEvent InvalidEventID(Value)
    End If

End Property

Public Property Get EventID() As String
    EventID = m_EventID
End Property

Public Property Let StartDate(Value As Date)
    If varType(Value) = vbDate Then
        m_StartDate = Value
    Else
        RaiseEvent InvalidStartDate(Value)
    End If
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let Observer(Value As String)
    If varType(Value) = vbString Then
        m_Observer = Value
    Else
        RaiseEvent InvalidObserver(Value)
    End If
End Property

Public Property Get Observer() As String
    Observer = m_Observer
End Property

Public Property Let Comments(Value As String)
    If varType(Value) = vbString Then
        m_Comments = Value
    Else
        RaiseEvent InvalidComments(Value)
    End If
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
            "Error encountered (#" & Err.Number & " - Class_Initialize[SamplingEvent class])"
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
            "Error encountered (#" & Err.Number & " - Class_Terminate[SamplingEvent class])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          UpdateStartDate
' Description:  Save event data to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/17/2017 - for NCPN tools
' Revisions:
'   BLC, 7/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub UpdateStartDate()
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "u_event_startdate"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "tblEvent"
        Params(1) = .EventID
        Params(2) = .StartDate
                
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateStartDate[SamplingEvent class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateObserver
' Description:  Save event data to database
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
    
    Template = "u_event_observer"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "tblEvent"
        Params(1) = .EventID
        Params(2) = .Observer
                
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateObserver[SamplingEvent class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateComments
' Description:  Save event data to database
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
    
    Template = "u_event_comments"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "tblEvent"
        Params(1) = .EventID
        Params(2) = .Comments
                
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateComments[SamplingEvent class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save event data to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 7/17/2017 - for NCPN tools
' Revisions:
'   BLC, 7/17/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_event_data"
    
    Dim Params(0 To 5) As Variant

    With Me
        Params(0) = "tblEvent"
        Params(1) = .EventID
        Params(2) = .StartDate
        Params(3) = .Observer
        Params(4) = .Comments
        
        If IsUpdate Then
            Template = "u_event_data"
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
                "Error encounter (#" & Err.Number & " - SaveToDb[SamplingEvent class])"
    End Select
    Resume Exit_Handler
End Sub