Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        InvasiveCoverSpecies
' Level:        Application class
' Version:      1.04
'
' Description:  Invasive cover species object related properties, events, functions & procedures for UI display
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 4/17/2017
' References:   -
' Revisions:    BLC - 4/17/2017 - 1.00 - initial version, adapted from Big Rivers UnderstoryCoverSpecies
'               BLC - 7/24/2017 - 1.01 - revised all percent cover properties to Single to
'                                        accommodate 0.5 (trace) percent covers
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
Private m_CoverSpecies As New CoverSpecies

'CoverSpecies includes:
'Private m_Species As New Species
'Private m_PctCover As Single
'Private m_QuadratID As Long
'Species includes:
'Private m_LUcode As String

Private m_IsDead As Byte
Private m_AverageCover As Single
Private m_PctCoverQ1 As Single
Private m_PctCoverQ2 As Single
Private m_PctCoverQ3 As Single
Private m_Position As Integer

Private m_SpeciesCoverID As Long    'species cover record ID

'---------------------
' Events
'---------------------
Public Event InvalidIsDead(Value As Byte)
Public Event InvalidAverageCover(Value As Single)
Public Event InvalidPctCoverQ1(Value As Single)
Public Event InvalidPctCoverQ2(Value As Single)
Public Event InvalidPctCoverQ3(Value As Single)
Public Event InvalidPosition(Value As Integer)

'-- base events (coverspecies)
Public Event InvalidQuadratID(Value As String)
Public Event InvalidPctCover(Value As Single) 'Integer)

'-- base events (species) --
Public Event InvalidMasterPlantCode(Value As String)
Public Event InvalidLUCode(Value As String)
Public Event InvalidFamily(Value As String)
Public Event InvalidSpecies(Value As String)
Public Event InvalidCode(Value As String)

'---------------------
' Properties
'---------------------
Public Property Let IsDead(Value As Byte)
    If varType(Value) = vbByte Then
        m_IsDead = Value
    Else
        RaiseEvent InvalidIsDead(Value)
    End If
End Property

Public Property Get IsDead() As Byte
    IsDead = m_IsDead
End Property

Public Property Let Position(Value As Integer)
    If varType(Value) = vbInteger Then
        m_Position = Value
    Else
        RaiseEvent InvalidPosition(Value)
    End If
End Property

Public Property Get Position() As Integer
    Position = m_Position
End Property

Public Property Let AverageCover(Value As Single)
    If varType(Value) = vbSingle Then
        m_AverageCover = Value
    Else
        RaiseEvent InvalidAverageCover(Value)
    End If
End Property

Public Property Get AverageCover() As Single
    AverageCover = m_AverageCover
End Property

Public Property Let PctCoverQ1(Value As Single)
    If IsBetween(Value, 0, 100, True) Then
        PctCoverQ1 = Value
    Else
        RaiseEvent InvalidPctCoverQ1(Value)
    End If
End Property

Public Property Get PctCoverQ1() As Single
    PctCoverQ1 = PctCoverQ1
End Property

Public Property Let PctCoverQ2(Value As Single)
    If IsBetween(Value, 0, 100, True) Then
        PctCoverQ2 = Value
    Else
        RaiseEvent InvalidPctCoverQ2(Value)
    End If
End Property

Public Property Get PctCoverQ2() As Single
    PctCoverQ2 = PctCoverQ2
End Property

Public Property Let PctCoverQ3(Value As Single)
    If IsBetween(Value, 0, 100, True) Then
        PctCoverQ3 = Value
    Else
        RaiseEvent InvalidPctCoverQ3(Value)
    End If
End Property

Public Property Get PctCoverQ3() As Single
    PctCoverQ3 = PctCoverQ3
End Property

Public Property Let SpeciesCoverID(Value As Long)
    m_SpeciesCoverID = Value
End Property

Public Property Get SpeciesCoverID() As Long
    SpeciesCoverID = m_SpeciesCoverID
End Property

' ---------------------------
' -- base class properties --
' ---------------------------
' NOTE: required since VBA does not support direct inheritance
'       or polymorphism like other OOP languages
' ---------------------------
' base class = Cover Species
' ---------------------------
Public Property Let QuadratID(Value As Long)
    m_CoverSpecies.QuadratID = Value
End Property

Public Property Get QuadratID() As Long
    QuadratID = m_CoverSpecies.QuadratID
End Property

Public Property Let PctCover(Value As Single) 'Integer)
    If IsBetween(Value, 0, 100, True) Then
        m_CoverSpecies.PctCover = Value
    Else
        RaiseEvent InvalidPctCover(Value)
    End If
End Property

Public Property Get PctCover() As Single 'Integer
    PctCover = m_CoverSpecies.PctCover
End Property

' ---------------------------
' base class = Species
' ---------------------------
Public Property Let ID(Value As Long)
    m_CoverSpecies.ID = Value
End Property

Public Property Get ID() As Long
    ID = m_CoverSpecies.ID
End Property

Public Property Let MasterPlantCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.MasterPlantCode = Value
    Else
        RaiseEvent InvalidMasterPlantCode(Value)
    End If
End Property

Public Property Get MasterPlantCode() As String
    MasterPlantCode = m_CoverSpecies.MasterPlantCode
End Property

Public Property Let COfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.COfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get COfamily() As String
    COfamily = m_CoverSpecies.COfamily
End Property

Public Property Let UTfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.UTfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get UTfamily() As String
    UTfamily = m_CoverSpecies.UTfamily
End Property

Public Property Let WYfamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.WYfamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get WYfamily() As String
    WYfamily = m_CoverSpecies.WYfamily
End Property

Public Property Let COspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.COspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get COspecies() As String
    COspecies = m_CoverSpecies.COspecies
End Property

Public Property Let UTspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.UTspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get UTspecies() As String
    UTspecies = m_CoverSpecies.UTspecies
End Property

Public Property Let WYspecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.WYspecies = Value
    Else
        RaiseEvent InvalidSpecies(Value)
    End If
End Property

Public Property Get WYspecies() As String
    WYspecies = m_CoverSpecies.WYspecies
End Property

Public Property Let LUCode(Value As String)
    'valid length varchar(25) but 6-letter lookup
    If Not IsNull(Value) And IsBetween(Len(Value), 1, 6, True) Then
        m_CoverSpecies.LUCode = Value
    Else
        RaiseEvent InvalidLUCode(Value)
    End If
End Property

Public Property Get LUCode() As String
    LUCode = m_CoverSpecies.LUCode
End Property

Public Property Let MasterFamily(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterFamily = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterFamily() As String
    MasterFamily = m_CoverSpecies.MasterFamily
End Property

Public Property Let MasterCode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.MasterCode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCode() As String
    MasterCode = m_CoverSpecies.MasterCode
End Property

Public Property Let MasterSpecies(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterSpecies = Value
    Else
        RaiseEvent InvalidFamily(Value)
    End If
End Property

Public Property Get MasterSpecies() As String
    MasterSpecies = m_CoverSpecies.MasterSpecies
End Property

Public Property Let UTcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.UTcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get UTcode() As String
    UTcode = m_CoverSpecies.UTcode
End Property

Public Property Let COcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.COcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get COcode() As String
    COcode = m_CoverSpecies.COcode
End Property

Public Property Let WYcode(Value As String)
    'valid length varchar(20) or ZLS
    If IsBetween(Len(Value), 1, 20, True) Then
        m_CoverSpecies.WYcode = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get WYcode() As String
    WYcode = m_CoverSpecies.WYcode
End Property

Public Property Let MasterCommonName(Value As String)
    'valid length varchar(50) or ZLS
    If IsBetween(Len(Value), 1, 50, True) Then
        m_CoverSpecies.MasterCommonName = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get MasterCommonName() As String
    MasterCommonName = m_CoverSpecies.MasterCommonName
End Property

Public Property Let Lifeform(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Lifeform = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Lifeform() As String
    Lifeform = m_CoverSpecies.Lifeform
End Property

Public Property Let Duration(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Duration = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Duration() As String
    Duration = m_CoverSpecies.Duration
End Property

Public Property Let Nativity(Value As String)
    'valid length varchar(255) or ZLS
    If IsBetween(Len(Value), 1, 255, True) Then
        m_CoverSpecies.Nativity = Value
    Else
        RaiseEvent InvalidCode(Value)
    End If
End Property

Public Property Get Nativity() As String
    Nativity = m_CoverSpecies.Nativity
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

'    MsgBox "Initializing...", vbOKOnly
    
    Set m_CoverSpecies = New CoverSpecies

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Initialize[InvasiveCoverSpecies class])"
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
    
'    MsgBox "Terminating...", vbOKOnly
        
    Set m_CoverSpecies = Nothing

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - Class_Terminate[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'======== Custom Methods ===========
'---------------------------------------------------------------------------------------
' SUB:          Init
' Description:  Lookup understory cover species based on the lookup code
' Parameters:   luCode - species 6-character lookup code from NCPN master plants (string)
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'---------------------------------------------------------------------------------------
Public Sub Init(LUCode As String)
On Error GoTo Err_Handler
    
    m_CoverSpecies.Init (LUCode)

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - Init[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          SaveToDb
' Description:  Save cover species based to database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell
' Adapted:      Bonnie Campbell, 4/19/2016 - for NCPN tools
' Revisions:
'   BLC, 4/19/2016 - initial version
'   BLC, 6/11/2016 - revised to GetTemplate()
'   BLC, 8/8/2016 - added update parameter to identify if this is an update vs. an insert
'   BLC, 7/18/2017 - code cleanup
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_invasive_cover_species"
    
    Dim Params(0 To 7) As Variant

    With Me
        Params(0) = "InvasiveCoverSpecies"
        Params(1) = .QuadratID
        Params(2) = .MasterPlantCode
        Params(3) = .PctCover
        Params(4) = .IsDead
        Params(5) = .Position
                
        If IsUpdate Then
            Template = "u_invasive_cover_species"
'            Params(6) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          AddSpeciesCover
' Description:  Add species cover record in database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   MarkK, November 17, 2011
'   https://www.access-programmers.co.uk/forums/showthread.php?t=218298
' Source/Date:  Bonnie Campbell, 7/18/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 7/18/2017 - initial version
'   BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'---------------------------------------------------------------------------------------
Public Sub AddSpeciesCover()
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_speciescover"
    
    Dim Params(0 To 5) As Variant

    With Me
        Params(0) = "SpeciesCover"
        Params(1) = .QuadratID
        Params(2) = .LUCode
        Params(3) = .IsDead
        Params(4) = .PctCover
        
        .ID = SetRecord(Template, Params)
        
        'retrieve the ID (requires MAX() vs. SELECT @@Identity
        'since each CurrentDb is a new db object & won't see the last insert
        With CurrDb
            Me.SpeciesCoverID = .OpenRecordset("SELECT MAX(ID) FROM SpeciesCover").Fields(0)
        End With
End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - AddSpeciesCover[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          UpdateSpeciesCover
' Description:  Update species cover record in database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:   -
' Source/Date:  Bonnie Campbell, 7/18/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 7/18/2017 - initial version
'---------------------------------------------------------------------------------------
Public Sub UpdateSpeciesCover()
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "u_speciescover"
    
    Dim Params(0 To 6) As Variant

    With Me
        Params(0) = "SpeciesCover"
        Params(1) = .SpeciesCoverID
        Params(2) = .QuadratID
        Params(3) = .LUCode
        Params(4) = .IsDead
        Params(5) = .PctCover
                
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - UpdateSpeciesCover[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub

'---------------------------------------------------------------------------------------
' SUB:          DeleteSpeciesCover
' Description:  Delete species cover record from database
' Parameters:   -
' Returns:      -
' Throws:       -
' References:
'   MarkK, November 17, 2011
'   https://www.access-programmers.co.uk/forums/showthread.php?t=218298
' Source/Date:  Bonnie Campbell, 7/18/2017 - for NCPN tools
' Adapted:      -
' Revisions:
'   BLC, 7/18/2017 - initial version
'   BLC, 10/4/2017 - switched CurrentDb to CurrDb property to avoid
'                     multiple open connections
'---------------------------------------------------------------------------------------
Public Sub DeleteSpeciesCover()
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "d_speciescover"
    
    Dim Params(0 To 2) As Variant

    With Me
        Params(0) = "SpeciesCover"
        Params(1) = .SpeciesCoverID
'        params(2) = .QuadratID
'        params(3) = .LUcode
'        params(4) = .IsDead
                
        .ID = SetRecord(Template, Params)
        
        'retrieve the species cover ID for the last inserted record
        'cannot use @@Identity here since it requires using same CurrentDb object
        '@ time CurrentDb is called, that's a new one >> use MAX(ID) instead
'        With CurrDb
'            Me.SpeciesCoverID = .OpenRecordset("SELECT MAX(ID) FROM SpeciesCover").Fields(0)
'        End With
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - DeleteSpeciesCover[InvasiveCoverSpecies class])"
    End Select
    Resume Exit_Handler
End Sub