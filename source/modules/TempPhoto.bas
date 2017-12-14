Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        TempPhoto
' Level:        Framework class
' Version:      1.00
'
' Description:  Temporary photo object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 12/12/2017
' References:   -
' Revisions:    BLC - 12/12/2017 - 1.00 - initial version
' =================================

'    [ID] [smallint] IDENTITY(1,1) NOT NULL,
'    [TempPhotographerID] [int] NULL,
'    [TempPhotoType] [nvarchar](2) NOT NULL,
'    [DigitalFilename] [nvarchar](15) NOT NULL,
'    [TakenDate] [datetime] NOT NULL,

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As Long
Private m_PhotoDate As Date
Private m_PhotoType As String '2
Private m_PhotographerID As Long
Private m_Path As String
Private m_Filename As String '10

Private m_Comments As AppComment

'---------------------
' Events
'---------------------
Public Event InvalidPhotoType(Value As String)
Public Event InvalidFilename(Value As String)
Public Event InvalidPath(Value As String)
Public Event InvalidPhotographerID(Value As Long)

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

Public Property Let PhotoDate(Value As Date)
    m_PhotoDate = Value
End Property

Public Property Get PhotoDate() As Date
    PhotoDate = m_PhotoDate
End Property

Public Property Let PhotoType(Value As String)
    Dim aryTypes() As String
    aryTypes = Split(PHOTO_TYPES, ",")
    If IsInArray(Value, aryTypes) Then
        m_PhotoType = Value
    Else
        RaiseEvent InvalidPhotoType(Value)
    End If
End Property

Public Property Get PhotoType() As String
    PhotoType = m_PhotoType
End Property

Public Property Let PhotographerID(Value As Long)
    m_PhotographerID = Value
End Property

Public Property Get PhotographerID() As Long
    PhotographerID = m_PhotographerID
End Property
    
Public Property Let FileName(Value As String)
    If FileExists(Value) Then
        m_Filename = Value
    Else
        RaiseEvent InvalidFilename(Value)
    End If
End Property

Public Property Get FileName() As String
    FileName = m_Filename
End Property

Public Property Let Path(Value As String)
    If FolderExists(Path) Then
        m_Path = Value
    Else
        RaiseEvent InvalidPath(Value)
    End If
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

'---------------------
' Events
'---------------------

'---------------------------------------------------------------------------------------
' EVENT:        InvalidFilename
' Description:  Responds to invalid filenames
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, 12/12/2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 12/12/2017 - initial version
'---------------------------------------------------------------------------------------
Private Sub TempPhoto_OnInvalidFilenameEvent()
On Error GoTo Err_Handler

    'user message
    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & "Please check your " & _
                        "photo filename && retry if necessary. " & _
                        vbCrLf & "[" & Me.FileName & " - InvalidFilename]" & _
                        "|Type" & PARAM_SEPARATOR & "caution" & _
                        "|Title" & PARAM_SEPARATOR & "Invalid Filename!"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - TempPhoto_OnInvalidFilenameEvent[TempPhoto class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' EVENT:        InvalidPath
' Description:  Responds to invalid file paths
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/date:  Bonnie Campbell, 12/12/2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 12/12/2017 - initial version
'---------------------------------------------------------------------------------------
Private Sub TempPhoto_OnInvalidPathEvent()
On Error GoTo Err_Handler

    'user message
    DoCmd.OpenForm "MsgOverlay", acNormal, , , , acDialog, _
                "msg" & PARAM_SEPARATOR & "Please check your " & _
                        "photo path && retry if necessary. " & _
                        vbCrLf & "[" & Me.Path & " - InvalidPath]" & _
                        "|Type" & PARAM_SEPARATOR & "caution" & _
                        "|Title" & PARAM_SEPARATOR & "Invalid File Path!"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - TempPhoto_OnInvalidPathEvent[TempPhoto class])"
    End Select
    Resume Exit_Handler
End Sub
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
' Source/date:  Bonnie Campbell, 12/12/2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 12/12/2017 - initial version
' ---------------------------------
Private Sub Class_Initialize()
On Error GoTo Err_Handler

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Class_Initialize[TempPhoto class])"
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
' Source/date:  Bonnie Campbell, 12/12/2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 12/12/2017 - initial version
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[TempPhoto class])"
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
' Source/date:  Bonnie Campbell, 12/12/2016 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 12/12/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_usys_temp_photo"
    
    Dim Params(0 To 7) As Variant
    
    With Me
        Params(0) = "TempPhoto"
        Params(1) = .EventID
        Params(2) = .PhotographerID
        Params(3) = .FileName
        Params(4) = .Path
        Params(5) = .PhotoDate
        Params(6) = .PhotoType
              
        If IsUpdate Then
            Template = "u_usys_temp_photo"
            Params(7) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

    'set observer/recorder
'    SetObserverRecorder Me, "Photo"

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[TempPhoto class])"
    End Select
    Resume Exit_Handler
End Sub