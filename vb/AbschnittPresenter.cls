VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AbschnittPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private objFSO

Private Sub Class_Initialize()
Set objFSO = CreateObject("Scripting.FileSystemObject")
End Sub



Public Function getSubFolderNames(name As String)

    If name = "" Then
        Set objFolder = objFSO.GetFolder("C:\XRAY\sections\")
    Else
        Set objFolder = objFSO.GetFolder("C:\XRAY\sections\" & name & "\")
    End If
    
    Dim items(30) As String
    Dim i As Integer
    i = 0
    For Each objsubFolder In objFolder.subfolders
      items(i) = objsubFolder.name
      i = i + 1
    Next objsubFolder

getSubFolderNames = items
End Function
Public Function getFile(parentDirectory As String, childDirectory As String)
   
    Set objFolder = objFSO.GetFolder("C:\XRAY\sections\" & parentDirectory & "\" & childDirectory & "\")
    Dim items(30) As String
    Dim i As Integer
    i = 0
    For Each objFile In objFolder.Files
      items(i) = objFile.name
      i = i + 1
    Next objFile

getFile = items
End Function

Public Sub setItems(comboboxObect As ComboBox, items() As String)
comboboxObect.Clear
For i = 0 To UBound(items)
If (items(i) <> "") Then
 comboboxObect.AddItem items(i)
 End If
 Next i

End Sub

Public Function insertSection(grandParentDirectory As String, parentDirectory As String, filename As String) As Boolean
If (grandParentDirectory = "" Or grandParentDirectory = "" Or filename = "") Then
    MsgBox "insert valid filename"
    insertSection = False
Else
    'Selection.InsertBreak Type:=wdPageBreak
    Dim sectionPath As String
    sectionPath = "C:\XRAY\sections\" & grandParentDirectory & "\" & parentDirectory & "\" & filename
     Selection.InsertFile filename:=sectionPath, Range:="", _
        ConfirmConversions:=False, Link:=False, Attachment:=False
        insertSection = True
End If


End Function

