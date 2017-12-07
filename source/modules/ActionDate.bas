Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        ActionDate
' Level:        Framework class
' Version:      1.02
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
'               BLC - 10/6/2017  - 1.02 - removed GetClass() after Factory class instatiation implemented
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

Public Property Let Contact(Value As person)
    m_Contact = Value
End Property

Public Property Get Contact() As person
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
' handled by Factory class

'======== Standard Methods ===========