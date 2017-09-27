Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Implements IVariantComparator

' =================================
' CLASS:        CReverseComparator
' Level:        Framework class
' Version:      1.00
'
' Description:  Comparison object related properties, events, functions & procedures
'
' Instancing:   PublicNotCreatable
'               Class is accessible w/in enclosing project & projects that reference it
'               Instances of class can only be created by modules w/in the enclosing project.
'               Modules in other projects may reference class name as a declared type
'               but may not instantiate class using new or the CreateObject function.
'
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' References:
' Revisions:    --------------- Reference Library ------------------
'               BLC - 9/27/2017  - 1.00 - initial version
' =================================

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
' FUNCTION:     IVariantComparator_Compare
' Description:  Compares two variants for their sort order.
'
'               IVariantComparator provides a method, compare, that imposes a
'               total ordering over a collection of variants.
'               A class that implements IVariantComparator, called a Comparator,
'               can be passed to the Arrays.sort and Collections.sort methods
'               to precisely control the sort order of the elements.
'
'               This function should exhibit several necessary behaviors:
'                 1) compare(x,y)=-  compare(y,x) for all x,y
'                 2) compare(x,y)>= 0 for all x,y
'                 3) compare(x,y)>=0
'                    compare(y,z)>=0 implies compare(x,z)> 0 for all x,y,z
' Assumptions:  -
' Parameters:
'               -
' Returns:      -1 --> v1 should be sorted ahead of v2
'               +1 --> v2 should be sorted ahead of v1
'                0 --> the two objects are of equal precedence
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function IVariantComparator_Compare(ByRef v1 As Variant, ByRef v2 As Variant) As Long
On Error GoTo Err_Handler
    
    IVariantComparator_Compare = v2 - v1

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IVariantComparator_Compare[CReverseComparator class])"
    End Select
    Resume Exit_Handler
End Function

' ---------------------------------
' FUNCTION:     IVariantComparator_Reverse
' Description:  Compares two variants for their sort order.
'
' Assumptions:  -
' Parameters:
'               -
' Returns:      -1 --> v1 should be sorted ahead of v2
'               +1 --> v2 should be sorted ahead of v1
'                0 --> the two objects are of equal precedence
' Throws:       none
' References:   -
' Source/date:
'   Austin D., July 11, 2016
'   https://stackoverflow.com/questions/3587662/how-do-i-sort-a-collection
' Adapted:      Bonnie Campbell, September 27, 2017 - for NCPN tools
' Revisions:
'   BLC - 9/27/2017 - initial version
' ---------------------------------
Public Function IVariantComparator_Reverse(ByRef v1 As Variant, ByRef v2 As Variant) As Long
On Error GoTo Err_Handler
    
    IVariantComparator_Reverse = v2 - v1

Exit_Handler:
    'cleanup
    Exit Function
    
Err_Handler:
    Select Case Err.Number
      Case Else
        MsgBox "Error #" & Err.Number & ": " & Err.Description, vbCritical, _
            "Error encountered (#" & Err.Number & " - IVariantComparator_Reverse[CReverseComparator class])"
    End Select
    Resume Exit_Handler
End Function