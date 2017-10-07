Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        Feature
' Level:        Framework class
' Version:      1.04
'
' Description:  Feature object related properties, events, functions & procedures
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
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.02 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.03 - switched CurrentDb to CurrDb property to avoid
'                                       multiple open connections
'               BLC - 10/6/2017 - 1.04 - removed GetClass() after Factory class instatiation implemented after Factory class instatiation implemented
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_LocationID As Integer
Private m_Name As String
Private m_Description As String
Private m_Directions As String
Private m_Sequence As Integer

'---------------------
' Events
'---------------------
Public Event InvalidID()
Public Event InvalidName(Name As String)
Public Event InvalidDescription(Description As String)
Public Event InvalidDirections(Directions As String)
Public Event InvalidSequence(Sequence As Integer)

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
End Property

Public Property Get LocationID() As Long
   LocationID = m_LocationID
End Property

Public Property Let Name(Value As String)
    'feature names are 1 or 2 characters (letters only)
    If Len(Trim(Value)) < 3 And IsAlpha(Value) Then
        m_Name = Value
    Else
        RaiseEvent InvalidName(Value)
    End If
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Description(Value As String)
    'descriptions - 255 characters or less
    If Len(Trim(Value)) < 256 Then
        m_Description = Value
    Else
        RaiseEvent InvalidDescription(Value)
    End If
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Directions(Value As String)
    m_Directions = Value
End Property

Public Property Get Directions() As String
    Directions = m_Directions
End Property

Public Property Let Sequence(Value As Integer)
    If Value > -1 Then
        m_Sequence = Value
    Else
        RaiseEvent InvalidSequence(Value)
    End If
End Property

Public Property Get Sequence() As Integer
    Sequence = m_Sequence
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[Feature class])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[Feature class])"
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
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_feature"
    
    Dim Params(0 To 6) As Variant
    
    With Me
        Params(0) = "Feature"
        Params(1) = .LocationID
        Params(2) = .Name
        Params(3) = .Description
        Params(4) = .Directions
        
        If IsUpdate Then
            Template = "u_feature"
            Params(5) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[Feature class])"
    End Select
    Resume Exit_Handler
End Sub