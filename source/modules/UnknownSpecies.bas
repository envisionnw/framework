Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        UnknownSpecies
' Level:        Framework class
' Version:      1.10
'
' Description:  UnknownSpecies object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'       Private modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 10/28/2015
' References:   -
' Revisions:    BLC - 11/12/2017 - 1.00 - initial version
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long    'Long
Private m_UnknownCode As String  'Text(15)
Private m_PlantType As String 'Text(15)
Private m_PlantDescription As String 'Memo >> Text(255)
Private m_SalientFeature As String 'Text(255)
Private m_LeafType As String 'Text(50)
Private m_LeafMargin As String 'Text(50)
Private m_LeafCharacter As String 'Text(255)
Private m_StemCharacter As String 'Text(255)
Private m_FlowerCharacter As String 'Text(255)
Private m_GeneralCharacter As String 'Text(255)
Private m_ForbGrassType As String 'Text(10)
Private m_PerennialGrassType As String 'Text(15)

Private m_TransectPosition As Double 'Integer > Double??
Private m_LocationID As Long 'Long

Private m_BestGuess As String 'Text(50)
Private m_HasPhotos As Byte 'Byte
Private m_Collected As Byte 'Byte
Private m_CollectionMethod As String  'Text(50)
Private m_CollectedByID As Long 'Long

Private m_ConfirmedCode As String 'Text(50)
Private m_IdentifiedByID As Long 'Long
Private m_IdentifiedDate As Date 'DateTime

'---------------------
' Events
'---------------------
Public Event InvalidCharacteristics(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let ID(Value As Long)
    m_ID = Value
End Property

Public Property Get ID() As Long
    ID = m_ID
End Property

Public Property Let BestGuess(Value As String)
    m_BestGuess = Value
End Property

Public Property Get BestGuess() As String
    BestGuess = m_BestGuess
End Property

Public Property Let UnknownCode(Value As String)
    m_UnknownCode = Value
End Property

Public Property Get UnknownCode() As String
    UnknownCode = m_UnknownCode
End Property

Public Property Let PlantType(Value As String)
    m_PlantType = Value
End Property

Public Property Get PlantType() As String
    PlantType = m_PlantType
End Property

Public Property Let PlantDescription(Value As String)
    m_PlantDescription = Value
End Property

Public Property Get PlantDescription() As String
    PlantDescription = m_PlantDescription
End Property

Public Property Let SalientFeature(Value As String)
    m_SalientFeature = Value
End Property

Public Property Get SalientFeature() As String
    SalientFeature = m_SalientFeature
End Property

Public Property Let LeafType(Value As String)
    m_LeafType = Value
End Property

Public Property Get LeafType() As String
    LeafType = m_LeafType
End Property

Public Property Let LeafMargin(Value As String)
    m_LeafMargin = Value
End Property

Public Property Get LeafMargin() As String
    LeafMargin = m_LeafMargin
End Property

Public Property Let LeafCharacter(Value As String)
    m_LeafCharacter = Value
End Property

Public Property Get LeafCharacter() As String
    LeafCharacter = m_LeafCharacter
End Property

Public Property Let StemCharacter(Value As String)
    m_StemCharacter = Value
End Property

Public Property Get StemCharacter() As String
    StemCharacter = m_StemCharacter
End Property

Public Property Let FlowerCharacter(Value As String)
    m_FlowerCharacter = Value
End Property

Public Property Get FlowerCharacter() As String
    FlowerCharacter = m_FlowerCharacter
End Property

Public Property Let GeneralCharacter(Value As String)
    m_GeneralCharacter = Value
End Property

Public Property Get GeneralCharacter() As String
    GeneralCharacter = m_GeneralCharacter
End Property

Public Property Let ForbGrassType(Value As String)
    m_ForbGrassType = Value
End Property

Public Property Get ForbGrassType() As String
    ForbGrassType = m_ForbGrassType
End Property

Public Property Let PerennialGrassType(Value As String)
    m_PerennialGrassType = Value
End Property

Public Property Get PerennialGrassType() As String
    PerennialGrassType = m_PerennialGrassType
End Property

'unknown location
Public Property Let TransectPosition(Value As Double)
    m_TransectPosition = Value
End Property

Public Property Get TransectPosition() As Double
    TransectPosition = m_TransectPosition
End Property

Public Property Let LocationID(Value As Long)
    m_LocationID = Value
End Property

Public Property Get LocationID() As Long
    LocationID = m_LocationID
End Property


'collection info
Public Property Let Collected(Value As Byte)
    m_Collected = Value
End Property

Public Property Get Collected() As Byte
    Collected = m_Collected
End Property

Public Property Let HasPhotos(Value As Byte)
    m_HasPhotos = Value
End Property

Public Property Get HasPhotos() As Byte
    HasPhotos = m_HasPhotos
End Property

Public Property Let CollectionMethod(Value As String)
    m_CollectionMethod = Value
End Property

Public Property Get CollectionMethod() As String
    CollectionMethod = m_CollectionMethod
End Property

Public Property Let CollectedByID(Value As Long)
    m_CollectedByID = Value
End Property

Public Property Get CollectedByID() As Long
    CollectedByID = m_CollectedByID
End Property


'identification info
Public Property Let IdentifiedByID(Value As Long)
    m_IdentifiedByID = Value
End Property

Public Property Get IdentifiedByID() As Long
    IdentifiedByID = m_IdentifiedByID
End Property

Public Property Let ConfirmedCode(Value As String)
    m_ConfirmedCode = Value
End Property

Public Property Get ConfirmedCode() As String
    ConfirmedCode = m_ConfirmedCode
End Property

Public Property Let IdentifiedDate(Value As Date)
    m_IdentifiedDate = Value
End Property

Public Property Get IdentifiedDate() As Date
    IdentifiedDate = m_IdentifiedDate
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
' Source/date:  Bonnie Campbell, 11/12/2017 - for NCPN tools
' Adapted:      -
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[UnknownSpecies class])"
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
' Source/Date:  Bonnie Campbell, 11/12/2017 - for NCPN tools
' Adapted:      -
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[UnknownSpecies class])"
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
' Source/Date:  Bonnie Campbell, 11/12/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 11/12/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_unknown"
    
    Dim Params(0 To 26) As Variant

    With Me
        Params(0) = "UnknownSpecies"
        Params(1) = .UnknownCode
        Params(2) = .PlantType
        Params(3) = .PlantDescription
        Params(4) = .SalientFeature
        Params(5) = .LeafType
        Params(6) = .LeafMargin
        Params(7) = .LeafCharacter
        Params(8) = .StemCharacter
        Params(9) = .FlowerCharacter
        Params(10) = .GeneralCharacter
        Params(11) = .ForbGrassType
        Params(12) = .PerennialGrassType
        Params(13) = .BestGuess
        Params(14) = .HasPhotos
        Params(15) = .Collected
        Params(16) = .CollectionMethod
        Params(17) = .LocationID
        Params(18) = .CollectedByID
'        Params(19) =.ConfirmedCode
'        Params(20) =.IdentifiedDate
'        Params(21) = .IdentifiedByID
'        Params(22) = .TransectPosition
        
        If IsUpdate Then
            Template = "u_unknown"
            Params(25) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                        "Error encounter (#" & Err.Number & " - SaveToDb[UnknownSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          IdentifyUnknown
' Description:  -
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   Fionnuala, February 2, 2009
'   David W. Fenton, October 27, 2009
'   http://stackoverflow.com/questions/595132/how-to-get-id-of-newly-inserted-record-using-excel-vba
' Source/Date:  Bonnie Campbell, 11/12/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 11/12/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub IdentifyUnknown(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "u_unknown_identify"
    
    Dim Params(0 To 5) As Variant

    With Me
        Params(0) = "UnknownSpecies"
        Params(1) = .ID
        Params(2) = .ConfirmedCode
        Params(3) = .IdentifiedDate
        Params(4) = .IdentifiedByID
                
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                        "Error encounter (#" & Err.Number & " - IdentifyUnknown[UnknownSpecies class])"
    End Select
    Resume Exit_Handler
End Sub