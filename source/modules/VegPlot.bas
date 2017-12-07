Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        VegPlot
' Level:        Framework class
' Version:      1.14
'
' Description:  VegPlot object related properties, events, functions & procedures
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
'               BLC - 1/12/2017 - 1.02 - added % litter, woody debris (all parks)
'                                        replaced FilamentousAlgae (boolean) w/
'                                        PctFilamentousAlgae (%)
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.03 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC, 10/4/2017  - 1.04 - SaveToDb() code cleaup
'               BLC - 10/6/2017 - 1.05 - removed GetClass() after Factory class instatiation implemented
'               BLC - 10/31/2017 - 1.06 - add ReplicatePlot, CalibrationPlot flags
'               BLC - 11/2/2017 - 1.07 - changed % covers to double from integer, added MSS, WCC, ARC % cover
'               BLC - 11/3/2017 - 1.08 - add ModalSedimentSize_ID property
'               BLC - 11/8/2017 - 1.09 - add PctModalSedimentSize property
'               BLC - 11/10/2017 - 1.10 - revised to HasSocialTrails & PctFines (plurals)
'               BLC - 11/11/2017 - 1.11 - revised PercentWoodyDebris > PctWoodyDebris, added PctStandingDead
'               BLC - 11/12/2017 - 1.12 - revised booleans to byte for 1,0 values
'               BLC - 11/26/2017 - 1.13 - revised HasSocialTrails to PctSocialTrails
'               BLC - 12/5/2017  - 1.14 - added BeaverBrowse
' =================================

'---------------------
' Declarations
'---------------------
Private m_ID As Long
Private m_EventID As Long
Private m_SiteID As Long
Private m_FeatureID As Long
Private m_VegTransectID As Long
Private m_PlotNumber As Integer
Private m_PlotDistance As Integer
Private m_ModalSedimentSizeID As Long
Private m_ModalSedimentSize As String '3
Private m_PctFines As Double
Private m_PctWater As Double
Private m_UnderstoryRootedPctCover As Double
Private m_WoodyCanopyPctCover As Double
Private m_AllRootedPctCover As Double
Private m_PctFilamentousAlgae As Double
Private m_PctModalSedimentSize As Double
Private m_PctLitter As Double
Private m_PctWoodyDebris As Double
Private m_PctStandingDead As Double
Private m_PlotDensity As Integer
Private m_NoCanopyVeg As Byte
Private m_NoRootedVeg As Byte
Private m_HasSocialTrails As Byte
Private m_NoIndicatorSpecies As Byte
Private m_ReplicatePlot As Byte
Private m_CalibrationPlot As Byte
Private m_PctSocialTrails As Double
Private m_BeaverBrowse As Byte

'---------------------
' Events
'---------------------
Public Event InvalidSizeClass(Value As String)
Public Event InvalidPlotDensity(Value As Integer)
Public Event InvalidPercent(Value As Double)

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

Public Property Let SiteID(Value As Long)
    m_SiteID = Value
End Property

Public Property Get SiteID() As Long
    SiteID = m_SiteID
End Property

Public Property Let FeatureID(Value As Long)
    m_FeatureID = Value
End Property

Public Property Get FeatureID() As Long
    FeatureID = m_FeatureID
End Property

Public Property Let VegTransectID(Value As Long)
    m_VegTransectID = Value
End Property

Public Property Get VegTransectID() As Long
    VegTransectID = m_VegTransectID
End Property

Public Property Let PlotNumber(Value As Integer)
    m_PlotNumber = Value
End Property

Public Property Get PlotNumber() As Integer
    PlotNumber = m_PlotNumber
End Property

Public Property Let PlotDistance(Value As Integer)
    m_PlotDistance = Value
End Property

Public Property Get PlotDistance() As Integer
    PlotDistance = m_PlotDistance
End Property

Public Property Let ModalSedimentSizeID(Value As Long)
    m_ModalSedimentSizeID = Value
End Property

Public Property Get ModalSedimentSizeID() As Long
    ModalSedimentSizeID = m_ModalSedimentSizeID
End Property

Public Property Let ModalSedimentSize(Value As String)
    'determine if valid ModWentworthClassSize
    Dim i As Integer
    For i = ModWentworthClassSize.[_First] To ModWentworthClassSize.[_Last]
'        If (ModWentworthClassSize(i) = Value) Then
            m_ModalSedimentSize = Value
'            Exit For
'        End If
        'set ModalSedimentSizeID from set class (assume current year for SamplingYear)
        m_ModalSedimentSizeID = 1 'g_ModSedimentSizes(m_ModalSedimentSize)
    Next
    'catch invalid values
    If Len(m_ModalSedimentSize) = 0 Then RaiseEvent InvalidSizeClass(Value)
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = m_ModalSedimentSize
End Property

Public Property Let PctFines(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctFines = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctFines() As Double
    PctFines = m_PctFines
End Property

Public Property Let PctWater(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctWater = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctWater() As Double
    PctWater = m_PctWater
End Property

Public Property Let UnderstoryRootedPctCover(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_UnderstoryRootedPctCover = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get UnderstoryRootedPctCover() As Double
    UnderstoryRootedPctCover = m_UnderstoryRootedPctCover
End Property

Public Property Let WoodyCanopyPctCover(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_WoodyCanopyPctCover = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get WoodyCanopyPctCover() As Double
    WoodyCanopyPctCover = m_WoodyCanopyPctCover
End Property

Public Property Let AllRootedPctCover(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_AllRootedPctCover = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get AllRootedPctCover() As Double
    AllRootedPctCover = m_AllRootedPctCover
End Property

Public Property Let PctModalSedimentSize(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctModalSedimentSize = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctModalSedimentSize() As Double
    PctModalSedimentSize = m_PctModalSedimentSize
End Property

Public Property Let PctFilamentousAlgae(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctFilamentousAlgae = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctFilamentousAlgae() As Double
    PctFilamentousAlgae = m_PctFilamentousAlgae
End Property

Public Property Let PctLitter(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctLitter = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctLitter() As Double
    PctLitter = m_PctLitter
End Property

Public Property Let PctWoodyDebris(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctWoodyDebris = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctWoodyDebris() As Double
    PctWoodyDebris = m_PctWoodyDebris
End Property

Public Property Let PctStandingDead(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctStandingDead = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Get PctStandingDead() As Double
    PctStandingDead = m_PctStandingDead
End Property

Public Property Let PlotDensity(Value As Integer)
    Dim aryDensity() As String
    aryDensity = Split(PLOT_DENSITIES, ",")
    If IsInArray(CStr(Value), aryDensity) Then
        m_PlotDensity = CInt(Value)
    Else
        RaiseEvent InvalidPlotDensity(Value)
    End If
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = m_PlotDensity
End Property

Public Property Let NoCanopyVeg(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_NoCanopyVeg = Value
End Property

Public Property Get NoCanopyVeg() As Byte
    NoCanopyVeg = m_NoCanopyVeg
End Property

Public Property Let NoRootedVeg(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_NoRootedVeg = Value
End Property

Public Property Get NoRootedVeg() As Byte
    NoRootedVeg = m_NoRootedVeg
End Property

Public Property Let HasSocialTrails(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_HasSocialTrails = Value
End Property

Public Property Get HasSocialTrails() As Byte
    HasSocialTrails = m_HasSocialTrails
End Property

Public Property Let NoIndicatorSpecies(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_NoIndicatorSpecies = Value
End Property

Public Property Get NoIndicatorSpecies() As Byte
    NoIndicatorSpecies = m_NoIndicatorSpecies
End Property

Public Property Let CalibrationPlot(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_CalibrationPlot = Value
End Property

Public Property Get CalibrationPlot() As Byte
    CalibrationPlot = m_CalibrationPlot
End Property

Public Property Let ReplicatePlot(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_ReplicatePlot = Value
End Property

Public Property Get ReplicatePlot() As Byte
    ReplicatePlot = m_ReplicatePlot
End Property

Public Property Get PctSocialTrails() As Double
    PctSocialTrails = m_PctSocialTrails
End Property

Public Property Let PctSocialTrails(Value As Double)
    If IsBetween(Value, 0, 100, True) Then
        m_PctSocialTrails = Value
    Else
        RaiseEvent InvalidPercent(Value)
    End If
End Property

Public Property Let BeaverBrowse(Value As Byte)
    If Value = 1 Or Value = 0 Then _
        m_BeaverBrowse = Value
End Property

Public Property Get BeaverBrowse() As Byte
    BeaverBrowse = m_BeaverBrowse
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
                "Error encounter (#" & Err.Number & " - Class_Initialize[VegPlot class])"
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
                "Error encounter (#" & Err.Number & " - Class_Terminate[VegPlot class])"
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
'   BLC, 1/12/2017 - added % litter, woody debris
'   BLC, 10/31/2017 - added ReplicatePlot, CalibrationPlot flags
'   BLC, 11/2/2017 - added % WCC, ARC, MSS
'   BLC, 11/8/2017 - added % MSS
'   BLC, 11/11/2017 - revised Percent > Pct (WoodyDebris, Litter, Fines, Water)
'   BLC, 12/5/2017 - added BeaverBrowse
'---------------------------------------------------------------------------------------
Public Sub SaveToDb(Optional IsUpdate As Boolean = False)
On Error GoTo Err_Handler
    
    Dim Template As String
    
    Template = "i_vegplot"
    
    Dim Params(0 To 28) As Variant

    With Me
        Params(0) = "VegPlot"
        Params(1) = .EventID
        Params(2) = .SiteID
        Params(3) = .FeatureID
        Params(4) = .VegTransectID
        Params(5) = .PlotNumber
        Params(6) = .PlotDistance
        Params(7) = .ModalSedimentSizeID    'vs. ModalSedimentSize class
        Params(8) = .PctFines
        Params(9) = .PctWater
        Params(10) = .UnderstoryRootedPctCover
        Params(11) = .WoodyCanopyPctCover
        Params(12) = .AllRootedPctCover
        Params(13) = .PctModalSedimentSize
        Params(14) = .PctFilamentousAlgae
        Params(15) = .PctLitter
        Params(16) = .PctWoodyDebris
        Params(17) = .PctStandingDead
        Params(18) = .PlotDensity
        Params(19) = .NoCanopyVeg
        Params(20) = .NoRootedVeg
        'Params(21) = .HasSocialTrails
        Params(22) = .NoIndicatorSpecies
        Params(23) = .CalibrationPlot
        Params(24) = .ReplicatePlot
        Params(25) = .PctSocialTrails
        Params(26) = .BeaverBrowse
        
        If IsUpdate Then
            Template = "u_vegplot"
            Params(28) = .ID
        End If
        
        .ID = SetRecord(Template, Params)
    End With

Exit_Handler:
    Exit Sub

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - SaveToDb[VegPlot class])"
    End Select
    Resume Exit_Handler
End Sub