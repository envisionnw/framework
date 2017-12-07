Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        AppUser
' Level:        Framework class
' Version:      1.03
'
' Description:  Application User object related properties, events, functions & procedures
'
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 11/3/2015
' References:   -
' Revisions:    BLC - 11/3/2015 - 1.00 - initial version
'               BLC - 8/23/2016 - 1.01 - added Initialize, Terminate, SaveToDb methods
'               BLC - 9/21/2017  - 1.02 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/6/2017 - 1.03 - removed GetClass() after Factory class instatiation implemented
' =================================

'---------------------
' Declarations
'---------------------
Dim AppUser As New person

Private m_Username As String
Private m_Password As String
Private m_Logins As Integer
Private m_Activity As String

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------
Public Property Let UserName(Value As String)
    m_Username = Value
End Property

Public Property Get UserName() As String
    UserName = m_Username
End Property

Public Property Let Password(Value As String)
    m_Password = Value
End Property

Public Property Get Password() As String
    Password = m_Password
End Property

Public Property Let Logins(Value As Integer)
    m_Logins = Value
End Property

Public Property Get Logins() As Integer
    Logins = m_Logins
End Property

Public Property Let activity(Value As String)
    m_Activity = Value
End Property

Public Property Get activity() As String
    activity = m_Activity
End Property

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========
' handled by Factory class

'======== Standard Methods ===========

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
                "Error encounter (#" & Err.Number & " - Class_Initialize[AppUser class])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[AppUser class])"
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
'   MarkK, September 11, 2013
'   http://www.access-programmers.co.uk/forums/showthread.php?t=253284
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/4/2016 - for NCPN tools
' Revisions:
'   BLC, 8/23/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_login"
    
    Dim Params() As Variant
    
    'dimension for contact
    ReDim Params(0 To 2) As Variant

    With Me
        Params(0) = "i_login" '"tsys_Db_Templates"
        Params(1) = .UserName
        Params(2) = .activity

'        If IsUpdate Then
'            template = "u_contact"
'            params(11) = .ID
'        End If
        
'        .ID = SetRecord(template, params)
        SetRecord Template, Params
    End With


Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[AppUser class])"
    End Select
    Resume Exit_Handler
End Sub