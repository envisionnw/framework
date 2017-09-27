Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

' =================================
' CLASS:        Factory
' Level:        Framework class
' Version:      1.00
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
' =================================

'---------------------
' Class Objects Available:

'---------------------


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
' FUNCTION:     New
' Description:  Creates new class object
'
' Assumptions:  -
' Parameters:   ClassName - name of class being created (string)
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