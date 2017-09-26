Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        ActionDate
' Level:        Framework class
' Version:      1.01
'
' Description:  Action date object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, November 10, 2015
' References:   -
' Revisions:    BLC - 11/10/2015 - 1.00 - initial version
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.01 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_FirstName As String
Private m_LastName As String
Private m_Name As String
Private m_Email As String
Private m_Role As String
Private m_Record As String
Private m_Contact As String
Private m_DateValue As Date
Private m_ActionType As String

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let FirstName(Value As String)
    m_FirstName = Value
End Property

Public Property Get FirstName() As String
    FirstName = m_FirstName
End Property

Public Property Let Record(Value As String)
    m_Record = Value
End Property

Public Property Get Record() As String
    Record = m_Record
End Property

Public Property Let Contact(Value As Person)
    m_Contact = Value
End Property

Public Property Get Contact() As Person
    Contact = m_Contact
End Property

Public Property Let DateValue(Value As Date)
    m_DateValue = Value
End Property

Public Property Get DateValue() As Date
    DateValue = m_DateValue
End Property

Public Property Let ActionType(Value As String)
    Select Case Value
        Case "Sample"
        Case "DataEntry"
        Case "Verification"
        Case "Download"
        Case "Change"
    End Select
    m_ActionType = Value
End Property

Public Property Get ActionType() As String
    ActionType = m_ActionType
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========

' ---------------------------------
' SUB:          GetClass
' Description:  Retrieve a new instance of the class
'               --------------------------------------------------------------------------
'               Classes in a library with PublicNotCreateable instancing cannot
'               create items of the class in other projects (using the New keyword)
'               Variables can be declared, but the class object isn't created
'
'               This function allows other projects to create new instances of the class object
'               In referencing projects, set a reference to this project & call the GetClass()
'               function to create the new class object:
'                   Dim NewActionDate as framework.ActionDate
'                   Set NewActionDate = framework.GetClass()
'               --------------------------------------------------------------------------
' Assumptions:  -
' Parameters:   -
' Returns:      New instance of the class
' Throws:       none
' References:
'   Chip Pearson, November 6, 2013
'   http://www.cpearson.com/excel/classes.aspx
' Source/date:  -
' Adapted:      Bonnie Campbell, September 21, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/21/2016 - initial version
' ---------------------------------
Public Function GetClass() As ActionDate
On Error GoTo Err_Handler

    Set GetClass = New ActionDate

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetClass[ActionDate class])"
    End Select
    Resume Exit_Handler
End Function

'======== Standard Methods ===========