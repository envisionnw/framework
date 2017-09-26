Attribute VB_Name = "Temp"
Option Compare Database
Option Explicit

Public Function AddDbClass(cls As String)
' Pelegrinus, April 7, 2012
'https://www.experts-exchange.com/questions/27664701/Use-VBA-Code-to-Save-class-modules-imported-into-Access.html

    Dim dir As String, ifile As String
    dir = "Z:\_____LIB\dev\git_projects\framework\modules"
    ifile = dir & "\" & cls & ".cls"
    
    Application.VBE.ActiveVBProject.VBComponents.Import ifile
    
    'Application.DoCmd.Save acModule, ifile
        
    'Application.RefreshDatabaseWindow

End Function
