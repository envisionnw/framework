Attribute VB_Name = "mod_Blank"
Option Compare Database
Option Explicit

' =================================
' MODULE:       mod_Utilities
' Level:        Framework module
' Version:      1.00
' Description:  Utility functions & procedures
'
' Source/date:  Bonnie Campbell, 9/14/2017
' Revisions:    BLC, 9/14/2017 - 1.00 - initial version
' =================================

' ---------------------------------
'  Declarations
' ---------------------------------

' ---------------------------------
'  Properties
' ---------------------------------

' ---------------------------------
'  Methods
' ---------------------------------
' ---------------------------------
' SUB:          CreateAppTempImages
' Description:  Create the temporary folder & files for the application
' Assumptions:  -
' Parameters:   -
' Returns:      -
' Throws:       none
' References:   none
' Requires:     -
' Source/date:  -
' Adapted:      Bonnie Campbell, September 14, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/14/2017  - initial version
' ---------------------------------
Public Sub CreateAppTempImages()
On Error GoTo Err_Handler
    

Exit_Handler:
    Exit Sub
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - CreateAppTempImages[mod_Utilities])"
    End Select
    Resume Exit_Handler
End Sub


