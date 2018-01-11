Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        ImportVegPlot
' Level:        Framework class
' Version:      1.02
'
' Description:  Import veg plot object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:  Bonnie Campbell, 2016
' References:   -
'   Comintern, November 2, 2016
'   http://stackoverflow.com/questions/40386553/long-to-wide-and-duplicate-column-when-row-has-data
' Revisions:    BLC - 2016       - 1.00 - initial version
'               --------------- Reference Library ------------------
'               BLC - 9/21/2017  - 1.01 - set class Instancing 2-PublicNotCreatable (VB_PredeclaredId = True),
'                                         VB_Exposed=True, added Property VarDescriptions, added GetClass() method
'               BLC - 10/6/2017  - 1.02 - removed GetClass() after Factory class instatiation implemented after Factory class instatiation implemented
' =================================

'Plot.cls

'---------------------
' Declarations
'---------------------
'---------------------
' Events
'---------------------

Private Type PlotData
    PlotID As Long
    VisitDate As Date
    LocationID As Long
    ModalSedimentSize As String
    PercentFine As Double
    PercentWater As Double
    UnderstoryRootedPctCover As Double
    PlotDensity As Integer
    NoCanopyVeg As Byte
    NoRootedVeg As Byte
    HasSocialTrails As Byte
    FilamentousAlgae As Byte
    NoIndicatorSpecies As Byte
    Litter As Double
End Type

Private this As PlotData
Private mCover As Scripting.Dictionary

Public Property Get PlotID() As Long
    PlotID = this.PlotID
End Property

Public Property Let PlotID(Value As Long)
    this.PlotID = Value
End Property

Public Property Get VisitDate() As Date
    VisitDate = this.VisitDate
End Property

Public Property Let VisitDate(Value As Date)
    this.VisitDate = Value
End Property

Public Property Get LocationID() As Long
    LocationID = this.LocationID
End Property

Public Property Let LocationID(Value As Long)
    this.LocationID = Value
End Property

Public Property Get ModalSedimentSize() As String
    ModalSedimentSize = this.ModalSedimentSize
End Property

Public Property Let ModalSedimentSize(Value As String)
    this.ModalSedimentSize = Value
End Property

Public Property Get PercentFine() As Double
    PercentFine = this.PercentFine
End Property

Public Property Let PercentFine(Value As Double)
    this.PercentFine = Value
End Property

Public Property Get PercentWater() As Double
    PercentWater = this.PercentWater
End Property

Public Property Let PercentWater(Value As Double)
    this.PercentWater = Value
End Property

Public Property Get UnderstoryRootedPctCover() As Double
    UnderstoryRootedPctCover = this.UnderstoryRootedPctCover
End Property

Public Property Let UnderstoryRootedPctCover(Value As Double)
    this.UnderstoryRootedPctCover = Value
End Property

Public Property Get PlotDensity() As Integer
    PlotDensity = this.PlotDensity
End Property

Public Property Let PlotDensity(Value As Integer)
    this.PlotDensity = Value
End Property

Public Property Get NoCanopyVeg() As Byte
    NoCanopyVeg = this.NoCanopyVeg
End Property

Public Property Let NoCanopyVeg(Value As Byte)
    this.NoCanopyVeg = Value
End Property

Public Property Get NoRootedVeg() As Byte
    NoRootedVeg = this.NoRootedVeg
End Property

Public Property Let NoRootedVeg(Value As Byte)
    this.NoRootedVeg = Value
End Property

Public Property Get HasSocialTrails() As Byte
    HasSocialTrails = this.HasSocialTrails
End Property

Public Property Let HasSocialTrails(Value As Byte)
    this.HasSocialTrails = Value
End Property

Public Property Get FilamentousAlgae() As Double
    FilamentousAlgae = this.FilamentousAlgae
End Property

Public Property Let FilamentousAlgae(Value As Double)
    this.FilamentousAlgae = Value
End Property

Public Property Get NoIndicatorSpecies() As Byte
    NoIndicatorSpecies = this.NoIndicatorSpecies
End Property

Public Property Let NoIndicatorSpecies(Value As Byte)
    this.NoIndicatorSpecies = Value
End Property

Public Property Get Litter() As Double
    Litter = this.Litter
End Property

Public Property Let Litter(Value As Double)
    this.Litter = Value
End Property

'Also in Plot.cls
Public Property Get CsvRows() As String
    Dim Key As Variant
    Dim output() As String
    ReDim output(mCover.Count - 1)
    Dim i As Long
    For Each Key In mCover.Keys
        Dim Temp(16) As String
        Temp(0) = this.PlotID
        Temp(1) = this.VisitDate
        Temp(2) = this.LocationID
        Temp(3) = this.ModalSedimentSize
        Temp(4) = this.PercentFine
        Temp(5) = this.PercentWater
        Temp(6) = this.UnderstoryRootedPctCover
        Temp(7) = this.PlotDensity
        Temp(8) = this.NoCanopyVeg
        Temp(9) = this.NoRootedVeg
        Temp(10) = this.HasSocialTrails
        Temp(11) = this.FilamentousAlgae
        Temp(12) = this.NoIndicatorSpecies
        Temp(13) = this.Litter
        Temp(14) = Key
        Temp(15) = mCover(Key)
        output(i) = Join(Temp, ",")
        i = i + 1
    Next Key
    CsvRows = Join(output, vbCrLf)
End Property

'Public Sub SampleUsage()
'    Dim plots As New Collection
'
'    With ActiveSheet
'        Dim col As Long
'        For col = 2 To 4
'            Dim current As Plot
'            Set current = New Plot
'            current.PlotId = .Cells(1, col).Value
'            current.DataDate = .Cells(2, col).Value
'            current.Location = .Cells(3, col).Value
'            Dim r As Long
'            For r = 4 To 6
'                Dim cover As String
'                cover = .Cells(r, col).Value
'                If cover <> vbNullString Then
'                    current.AddSpeciesCover .Cells(r, 1).Value, cover
'                End If
'            Next
'            plots.Add current
'        Next
'
'    End With
'
'    For Each current In plots
'        Debug.Print current.CsvRows
'    Next
'End Sub

'---------------------
' Methods
'---------------------

'======== Instancing Method ==========
' handled by Factory class

'======== Standard Methods ==========
Private Sub Class_Initialize()
    Set mCover = New Scripting.Dictionary
End Sub

Public Sub AddSpeciesCover(Species As String, cover As String)
    mCover.Add Species, cover
End Sub