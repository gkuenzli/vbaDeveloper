Attribute VB_Name = "VbaDeveloperConfig"
''
' VbaDeveloper Configuration Module
'
' This is a comment-only module aimed at configuring VbaDeveloper
' for this project.
'
'' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~ '


' ' Define the folder for export/import of code
' ' The path is relative to the folder containing this workbook
'
' ' You can include parameters in the path:
' '  - %ProjectName% : the name of the VBProject
' '  - %FileName% : the file name of the workbook (including extension)
'
' RelativeSourcePath = src\%FileName%



' ' Configure VbaDeveloper event-related behaviour
' ' requires that VbaDeveloper instantiate EventListner
'
' ImportAfterOpen = False
' FormatBeforeSave = False
' ExportAfterSave = True


' ' Discard some VBComponents from export/import operations
' ' Use as value the name of the component
' ' One line per ignored component
'
' Ignore = Tests
' Ignore = UserForm1
' Ignore = Sheet1
