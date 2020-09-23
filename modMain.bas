Attribute VB_Name = "modMain"
'------------------------------------------------------------------
' Name:           modMain (modMain.bas)
' Type:           Module
' Description:    -
'
' Author:         Koen Muilwijk
' Date:           10-1-2002
' E-mail:         deye_99@yahoo.com
' Copyright:      This work is copyrighted © 2001
'
' Comments:       -
'------------------------------------------------------------------
Option Explicit

'Load all gif images of a given directory in a ImageCombobox
Public Sub LoadFolderInIml(imlTarget As ImageList, strFolder As String)
Dim strFile As String, strKey As String
  
'  cmbTarget.ImageList = imlTarget
  strFile = Dir(strFolder & "*.gif")
  Do While strFile <> ""
    'Add the image to the image list
    strKey = Left(strFile, Len(strFile) - 4)
    imlTarget.ListImages.Add , strKey, LoadPicture(strFolder & strFile)
    
    'Ask for the next file
    strFile = Dir()
    
  Loop

End Sub

