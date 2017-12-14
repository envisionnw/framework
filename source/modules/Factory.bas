Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        Factory
' Level:        Framework class
' Version:      1.02
'
' Description:  Factory object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Note:         Factory is a Static Class
'               It cannot be instantiated or inherited (i.e. it's "static")
'               To create this static class the attribute below is set in the
'               exported class file & reimported:
'
'                   Attribute VB_PredeclaredId = True
'
'               This allows what is essentially a standard module to appear as if it were a class.
'
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' References:
' Revisions:    --------------- Reference Library ------------------
'               BLC - 9/27/2017  - 1.00 - initial version
'               BLC - 11/12/2017 - 1.01 - added unknown class
'               BLC - 12/14/2017 - 1.02 - added tempphoto class
' =================================

'---------------------------------------------------------------------------
' Class Objects Available:
'---------------------------------------------------------------------------
'   ActionDate              RootedSpecies
'   AppComment              SamplingEvent
'   AppUser                 Site
'   CoverSpecies            Species
'   CReverseComparator      Surface
'   EventVisit              SurfaceCover
'   ExifReader              SurveyFile
'   ExtArray                Tagline
'   Feature                 Task
'   ImportVegPlot           Template
'   InvasiveCoverSpecies    Transducer
'   IVariantComparator      Transect
'   Link                    UnderstoryCoverSpecies
'   Location                VegPlot
'   Observation             VegTransect
'   Park                    VegWalk
'   Person                  VegWalkSpecies
'   Photo                   Waterway
'   Quadrat                 WoodyCanopySpecies
'   RecordAction
'   TempPhoto               Unknown
'---------------------------------------------------------------------------

'---------------------
' Declarations
'---------------------

'---------------------
' Events
'---------------------

'---------------------
' Properties
'---------------------

'---------------------
' Methods
'---------------------
'======== Standard Methods ==========

' ---------------------------------
' FUNCTION:     NewActionDate
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewActionDate() As ActionDate
On Error GoTo Err_Handler
    
    Set NewActionDate = New ActionDate

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewActionDate[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewAppComment
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewAppComment() As AppComment
On Error GoTo Err_Handler
    
    Set NewAppComment = New AppComment

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewAppComment[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewAppUser
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewAppUser() As AppUser
On Error GoTo Err_Handler
    
    Set NewAppUser = New AppUser

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewAppUser[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewCoverSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewCoverSpecies() As CoverSpecies
On Error GoTo Err_Handler
    
    Set NewCoverSpecies = New CoverSpecies

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewCoverSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewCReverseComparator
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewCReverseComparator() As CReverseComparator
On Error GoTo Err_Handler
    
    Set NewCReverseComparator = New CReverseComparator

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewCReverseComparator[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewEventVisit
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewEventVisit() As EventVisit
On Error GoTo Err_Handler
    
    Set NewEventVisit = New EventVisit

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewEventVisit[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewExifReader
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewExifReader() As ExifReader
On Error GoTo Err_Handler
    
    Set NewExifReader = New ExifReader

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewExifReader[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewExtArray
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewExtArray()
On Error GoTo Err_Handler
    
    Set NewExtArray = New ExtArray

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewExtArray[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewFeature
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewFeature() As Feature
On Error GoTo Err_Handler
    
    Set NewFeature = New Feature

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewFeature[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewImportVegPlot
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewImportVegPlot() As ImportVegPlot
On Error GoTo Err_Handler
    
    Set NewImportVegPlot = New ImportVegPlot

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewImportVegPlot[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewInvasiveCoverSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewInvasiveCoverSpecies() As InvasiveCoverSpecies
On Error GoTo Err_Handler
    
    Set NewInvasiveCoverSpecies = New InvasiveCoverSpecies

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewInvasiveCoverSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewIVariantComparator
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewIVariantComparator() As IVariantComparator
On Error GoTo Err_Handler
    
    NewIVariantComparator = New IVariantComparator

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewIVariantComparator[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewLink
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewLink() As Link
On Error GoTo Err_Handler
    
    Set NewLink = New Link

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewLink[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewLocation
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewLocation() As Location
On Error GoTo Err_Handler
    
    Set NewLocation = New Location

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewLocation[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewObservation
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewObservation() As Observation
On Error GoTo Err_Handler
    
    Set NewObservation = New Observation

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewObservation[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewPark
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewPark() As Park
On Error GoTo Err_Handler
    
    Set NewPark = New Park
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewPark[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewPerson
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewPerson() As person
On Error GoTo Err_Handler
    
    Set NewPerson = New person

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewPerson[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewPhoto
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewPhoto() As Photo
On Error GoTo Err_Handler
    
    Set NewPhoto = New Photo

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewPhoto[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewQuadrat
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewQuadrat() As Quadrat
On Error GoTo Err_Handler
    
    Set NewQuadrat = New Quadrat

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewQuadrat[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewRecordAction
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewRecordAction() As RecordAction
On Error GoTo Err_Handler
    
    Set NewRecordAction = New RecordAction

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewRecordAction[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewRootedSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewRootedSpecies() As RootedSpecies
On Error GoTo Err_Handler
    
    Set NewRootedSpecies = New RootedSpecies

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewRootedSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSamplingEvent
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSamplingEvent() As SamplingEvent
On Error GoTo Err_Handler
    
    Set NewSamplingEvent = New SamplingEvent

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSamplingEvent[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSite
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSite() As Site
On Error GoTo Err_Handler
    
    Set NewSite = New Site

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSite[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSpecies() As Species
On Error GoTo Err_Handler
    
    Set NewSpecies = New Species
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSurface
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSurface() As Surface
On Error GoTo Err_Handler
    
    Set NewSurface = New Surface
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSurface[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSurfaceCover
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSurfaceCover() As SurfaceCover
On Error GoTo Err_Handler

    Set NewSurfaceCover = New SurfaceCover

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSurfaceCover[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewSurveyFile
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewSurveyFile() As SurveyFile
On Error GoTo Err_Handler
    
    Set NewSurveyFile = New SurveyFile

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewSurveyFile[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTagline
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewTagline() As Tagline
On Error GoTo Err_Handler
    
    Set NewTagline = New Tagline

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTagline[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTask
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewTask() As Task
On Error GoTo Err_Handler
    
    Set NewTask = New Task

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTask[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTemplate
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewTemplate() As Template
On Error GoTo Err_Handler
    
    Set NewTemplate = New Template

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTemplate[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTempPhoto
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 12/14/2017 - initial version
' ---------------------------------
Public Function NewTempPhoto() As TempPhoto
On Error GoTo Err_Handler
    
    Set NewTempPhoto = New TempPhoto

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTempPhoto[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTransducer
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewTransducer() As Transducer
On Error GoTo Err_Handler

    Set NewTransducer = New Transducer

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTransducer[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewTransect
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewTransect() As Transect
On Error GoTo Err_Handler
    
    Set NewTransect = New Transect

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewTransect[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewUnderstoryCoverSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewUnderstoryCoverSpecies() As UnderstoryCoverSpecies
On Error GoTo Err_Handler
    
    Set NewUnderstoryCoverSpecies = New UnderstoryCoverSpecies

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewUnderstoryCoverSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewUnknownSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, November 12, 2017 - for NCPN tools
' Revisions:
'   BLC - 11/12/2017 - initial version
' ---------------------------------
Public Function NewUnknownSpecies()
On Error GoTo Err_Handler
    
    Set NewUnknownSpecies = New UnknownSpecies

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewUnknownSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewVegPlot
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewVegPlot() As VegPlot
On Error GoTo Err_Handler
    
    Set NewVegPlot = New VegPlot

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewVegPlot[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewVegTransect
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewVegTransect() As VegTransect
On Error GoTo Err_Handler

    Set NewVegTransect = New VegTransect
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewVegTransect[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewVegWalk
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewVegWalk() As VegWalk
On Error GoTo Err_Handler
    
    Set NewVegWalk = New VegWalk
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewVegWalk[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewVegWalkSpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewVegWalkSpecies() As VegWalkSpecies
On Error GoTo Err_Handler
    
    Set NewVegWalkSpecies = New VegWalkSpecies
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewVegWalkSpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewWaterway
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewWaterway() As Waterway
On Error GoTo Err_Handler
    
    Set NewWaterway = New Waterway

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewWaterway[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     NewWoodyCanopySpecies
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewWoodyCanopySpecies() As WoodyCanopySpecies
On Error GoTo Err_Handler
    
    Set NewWoodyCanopySpecies = New WoodyCanopySpecies
    
Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - NewWoodyCanopySpecies[Factory class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     New
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   -
'               -
' Returns:      object of the desired class
' Throws:       none
' References:   -
' Source/date:
'   Hammond Mason, July 9, 2015
'   https://hammondmason.wordpress.com/2015/07/09/object-oriented-vba-design-patterns-simple-factory/
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function NewX()
On Error GoTo Err_Handler
    

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - New[Factory class])"
    End Select
    Resume Exit_Handler
End Function