Attribute VB_Name = "Stempel"
Sub add_Stempel()
    Dim picLoc As String
    Dim wrdPic As InlineShape
    Dim lngPercent2Scale As Long
    Dim lngOriginalHeight As Long
    Dim lngScaledHeight As Long
    Dim lngOriginalWidth As Long
    Dim lngScaledWidth As Long
    
    On Error GoTo errorHandling
    ' setting the location of the stempel
    Let picLoc = "C:\XRAY\img\" & "stempel.jpg"
    
    ' adding the picture to the selection area
    Set wrdPic = Selection.InlineShapes.AddPicture( _
        filename:=picLoc, LinkToFile:=False, SaveWithDocument:=True)
    wrdPic.Range.Font.Hidden = True
    ' to modify the height and the width independent
    wrdPic.LockAspectRatio = msoFalse
    
    'Here
        
    wrdPic.Range.Font.Hidden = False
    Exit Sub
    
errorHandling:
    Select Case Err.Number
    Case 5152
        ' to handle the case that someone has moved or deleted the stempel
        ' give the user the choice however to choose the picture hisself
        MsgBox Prompt:="Sorry, no Stempel is found within the default path" & vbCr & _
        "The default path : " & picLoc & vbCr & _
        "You might have moved or deleted it ? ", Buttons:=vbExclamation
        
     Case Else
        ' to handle if any other error has occured
        MsgBox Prompt:="Sorry, an error has occurred. Cannot continue." & _
                           vbCr & "(" & Err.Number & " - " & Err.Description & ")", _
                   Buttons:=vbCritical
    End Select
End Sub


