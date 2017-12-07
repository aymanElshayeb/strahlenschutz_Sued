Attribute VB_Name = "file_save"
Sub file_save()

    Dim dlgFileSaveAs As Dialog
    Dim path As String
    Dim lngAnswer As Long
    
    'Initialize the saveAs dialog
    Set dlgFileSaveAs = Dialogs(wdDialogFileSaveAs)
    
    'the chosen path must be specified hier
    path = "C:\XRAY\output"
    
    'to handle the case that someone has deleted the default saveAs folder
    On Error GoTo errorHandling
    ChDir (path)
    
    dlgFileSaveAs.name = path
    dlgFileSaveAs.Show
    Exit Sub
    
errorHandling:
    Select Case Err.Number
     Case 76
        ' to handle the case that someone has deleted the default saveAs folder
        ' give the user the choice however to choose a temporary path to save the document until the problem is resolved
        lngAnswer = MsgBox(Prompt:="Sorry, The default path for saving documents can't be found" & vbCr & _
        "The default path : " & path & vbCr & _
        "Do you want to enter a new path ? ", Buttons:=vbYesNo)
        If lngAnswer = vbYes Then
            path = create_newPath()
        End If
        
     Case Else
        ' to handle if any other error has occured
        MsgBox Prompt:="Sorry, an error has occurred. Cannot continue." & _
                           vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                   Buttons:=vbCritical
    End Select
     
End Sub

Private Function create_newPath() As String
            
    Dim tempDlgFileSaveAs As Dialog
    Dim lngDialog As Long
    Dim newPath As String
    
    ' to remind the user that the default path is still the same which is not working
   MsgBox Prompt:="The path will change only temporarily " & vbCr & _
   "Contact the adminstrator to set the new one as The default ", Buttons:=vbInformation
   
   'Initialize the temporary saveAs dialog
   Set tempDlgFileSaveAs = Dialogs(wdDialogFileSaveAs)
   tempDlgFileSaveAs.Show

    
End Function



