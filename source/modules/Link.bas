Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        Link
' Level:        Framework class
' Version:      1.01
'
' Description:  Link object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 10/30/2015
' References:
'  Maciej Los, April 5, 2011
'  http://www.codeproject.com/Questions/167323/Using-a-VS-Custom-Control-in-VBA-NOT-VB
' Revisions:    BLC - 10/30/2015 - 1.00 - initial version
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.01 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Integer
Private m_Text As String
Private m_Action As String
Private m_LinkFontColor As Long
Private m_LinkBgColor As Long
Private m_LinkVisible As Byte
Private m_LinkSeparatorVisible As Byte

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

Public Property Let text(Value As String)
    m_Text = Value
End Property

Public Property Get text() As String
    text = m_Text
End Property

Public Property Let action(Value As String)
    m_Action = Value
End Property

Public Property Get action() As String
    action = m_Action
End Property

Public Property Let LinkFontColor(Value As Long)
    m_LinkFontColor = Value
End Property

Public Property Get LinkFontColor() As Long
    LinkFontColor = m_LinkFontColor
End Property

Public Property Let LinkBgColor(Value As Long)
    If Len(Trim(Value)) < 0 Then Value = vbGreen '"#3F3F3F"
    m_LinkBgColor = Value
    
    'set font color to match
    Select Case Value
        Case vbGreen
            Me.LinkFontColor = vbBlack
        Case vbRed, vbBlue
            Me.LinkFontColor = vbWhite
    End Select
End Property

Public Property Get LinkBgColor() As Long
    LinkBgColor = m_LinkBgColor 'FormHeader.BackColor
End Property

Public Property Let LinkVisible(Value As Byte)
    m_LinkVisible = Value
End Property

Public Property Get LinkVisible() As Byte
    LinkVisible = m_LinkVisible
End Property

Public Property Let LinkSeparatorVisible(Value As Byte)
    m_LinkSeparatorVisible = Value
End Property

Public Property Get LinkSeparatorVisible() As Byte
    LinkSeparatorVisible = m_LinkSeparatorVisible
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
'                   Dim NewLink as framework.Link
'                   Set NewLink = framework.GetClass()
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
Public Function GetClass() As Link
On Error GoTo Err_Handler

    Set GetClass = New Link

Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetClass[Link class])"
    End Select
    Resume Exit_Handler
End Function

'======== Standard Methods ==========