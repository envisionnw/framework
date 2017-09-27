Option Compare Database
Option Explicit

'======== Instancing Method ==========

' ---------------------------------
' SUB:          GetClass
' Description:  Retrieve a new instance of the class
'               --------------------------------------------------------------------------
'               Classes in a library with PublicNotCreateable instancing cannot
'               create items of the class in other projects (using the New keyword)
'               Variables can be declared, but the class object isn't created
'
'               This function allows other projects to create new instances of the class object
'               In referencing projects, set a reference to this project & call the GetClass()
'               function to create the new class object:
'                   Dim NewActionDate as framework.ActionDate
'                   Set NewActionDate = framework.GetClass()
'               --------------------------------------------------------------------------
' Assumptions:  -
' Parameters:   ClassName - name of class to be instantiated (string)
' Returns:      New instance of the class
' Throws:       none
' References:
'   Chip Pearson, November 6, 2013
'   http://www.cpearson.com/excel/classes.aspx
' Source/date:  -
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2016 - initial version
' ---------------------------------
Public Function GetClass(ClassName As String) As Object
On Error GoTo Err_Handler

    Select Case ClassName
        Case ""
        Case ""
        Set GetClass = New ActionDate
    End Select
    
Exit_Handler:
    Exit Function

Err_Handler:
    Select Case Err.Number
        Case Else
            MsgBox "Error #" & Err.Description, vbCritical, _
                "Error encounter (#" & Err.Number & " - GetClass[ActionDate class])"
    End Select
    Resume Exit_Handler
End Function