Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        Tagline
' Level:        Framework class
' Version:      1.05
'
' Description:  Record Tagline object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 6/1/2016  - 1.01 - updated to use GetTemplate() in SaveToDb()
'               BLC - 8/8/2016  - 1.02 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.03 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.04 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC - 10/6/2017 - 1.05 - removed GetClass() after Factory class instatiation implemented
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_LineDistSource As String
Private m_LineDistSourceID As Long
Private m_LineDistType As String
Private m_LineDistance As Integer
Private m_HeightType As String
Private m_Height As Integer

'---------------------
' Events
'---------------------
Public Event InvalidLineDistSource(Value As String)
Public Event InvalidLineDistType(Value As String)
Public Event InvalidLineDistance(Value As Integer) 'in m
Public Event InvalidHeightType(Value As String)
Public Event InvalidHeight(Value As Integer)    'in cm

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let LineDistSource(Value As String)
    Dim arySources() As String
    arySources = Split(LINE_DIST_SOURCES, ",")
    If IsInArray(Value, arySources) Then
            m_LineDistSource = Value
    Else
        RaiseEvent InvalidLineDistSource(Value)
    End If
End Property

Public Property Get LineDistSource() As String
    LineDistSource = m_LineDistSource
End Property

Public Property Let LineDistSourceID(Value As Long)
    m_LineDistSourceID = Value
End Property

Public Property Get LineDistSourceID() As Long
    LineDistSourceID = m_LineDistSourceID
End Property

Public Property Let LineDistType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(LINE_DIST_TYPES, ",")
    If IsInArray(Value, aryTypes) Then
            m_LineDistType = Value
    Else
        RaiseEvent InvalidLineDistType(Value)
    End If
End Property

Public Property Get LineDistType() As String
    LineDistType = m_LineDistType
End Property

Public Property Let LineDistance(Value As Integer)
    m_LineDistance = Value
End Property

Public Property Get LineDistance() As Integer
    LineDistance = m_LineDistance
End Property

Public Property Let HeightType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(HEIGHT_TYPES, ",")
    If IsInArray(Value, aryTypes) Then
        m_HeightType = Value
    Else
        RaiseEvent InvalidHeightType(Value)
    End If
End Property

Public Property Get HeightType() As String
    HeightType = m_HeightType
End Property

Public Property Let Height(Value As Integer)
    m_Height = Value
End Property

Public Property Get Height() As Integer
    Height = m_Height
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[Tagline class])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[Tagline class])"
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
'   BLC, 6/1/2016 - updated to use GetTemplate()
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_tagline"
    
    Dim Params(0 To 8) As Variant

    With Me
        Params(0) = "Tagline"
        Params(1) = .LineDistSource
        Params(2) = .LineDistSourceID
        Params(3) = .LineDistType
        Params(4) = .LineDistance
        Params(5) = .HeightType
        Params(6) = .Height
        
        If IsUpdate Then
            Template = "u_tagline"
            Params(7) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With
    
'    'add a record for created by
'    Dim act As New RecordAction
'
'    With act
'        .RefAction = "R"
'        .ContactID = TempVars("UserID")
'        .RefID = Me.ID
'        .RefTable = "Tagline"
'        .SaveToDb
'    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[Tagline class])"
    End Select
    Resume Exit_Handler
End Sub