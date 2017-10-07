Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegWalk
' Level:        Framework class
' Version:      1.04
'
' Description:  Veg walk object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 4/19/2016
' References:   -
' Revisions:    BLC - 4/19/2016 - 1.00 - initial version
'               BLC - 8/8/2016  - 1.01 - SaveToDb() added update parameter to identify if
'                                        this is an update vs. an insert
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.02 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/4/2017 - 1.03 - SaveToDb() code cleanup
'               BLC - 10/6/2017 - 1.04 - removed GetClass() after Factory class instatiation implemented
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As Long
Private m_CollectionPlaceID As Long
Private m_CollectionType As String
Private m_StartDate As Date
Private m_CreateDate As Date
Private m_CreatedByID As Long
Private m_LastModified As Date
Private m_LastModifiedByID As Long

'---------------------
' Events
'---------------------
Public Event InvalidEventID(Value As Long)
Public Event InvalidCollectionPlaceID(Value As Long)
Public Event InvalidCollectionType(Value As String)
Public Event InvalidStartDate(Value As Date)
Public Event InvalidDate(Value As Date)
Public Event InvalidContactID(Value As Long)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let EventID(Value As Long)
    m_EventID = Value
End Property

Public Property Get EventID() As Long
    EventID = m_EventID
End Property

Public Property Let CollectionPlaceID(Value As Long)
    m_CollectionPlaceID = Value
End Property

Public Property Get CollectionPlaceID() As Long
    CollectionPlaceID = m_CollectionPlaceID
End Property

Public Property Let CollectionType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(COLLECTION_TYPES, ",")
    'check for valid collection type
    If IsInArray(Value, aryTypes) Then
        m_CollectionType = Value
    Else
        RaiseEvent InvalidCollectionType(Value)
    End If
End Property

Public Property Get CollectionType() As String
    CollectionType = m_CollectionType
End Property

Public Property Let StartDate(Value As Date)
    m_StartDate = Value
End Property

Public Property Get StartDate() As Date
    StartDate = m_StartDate
End Property

Public Property Let CreateDate(Value As Date)
    m_CreateDate = Value
End Property

Public Property Get CreateDate() As Date
    CreateDate = m_CreateDate
End Property

Public Property Let CreatedByID(Value As Long)
    m_CreatedByID = Value
End Property

Public Property Get CreatedByID() As Long
    CreatedByID = m_CreatedByID
End Property

Public Property Let LastModified(Value As Date)
    m_LastModified = Value
End Property

Public Property Get LastModified() As Date
    LastModified = m_LastModified
End Property

Public Property Let LastModifiedByID(Value As Long)
    m_LastModifiedByID = Value
End Property

Public Property Get LastModifiedByID() As Long
    LastModifiedByID = m_LastModifiedByID
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
'   BLC - 4/19/2016 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[VegWalk class])"
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
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'---------------------------------------------------------------------------------------
Private Sub Class_Terminate()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Terminate[VegWalk class])"
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
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegwalk"
    
    Dim Params(0 To 10) As Variant
    
    With Me
        Params(0) = "VegWalk"
        Params(1) = .EventID
        Params(2) = .CollectionPlaceID
        Params(3) = .CollectionType
        Params(4) = .StartDate
'        params(5) = .CreateDate
'        params(6) = .CreatedByID
'        params(7) = .LastModified
'        params(8) = .LastModifiedByID
        
        If IsUpdate Then
            Template = "u_vegwalk"
            Params(9) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[VegWalk class])"
    End Select
    Resume Exit_Handler
End Sub