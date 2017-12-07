Attribute VB_Name = "NeuerBerichtModule"

Sub openFileAndHideForm(name As String)
     NeuerBerichtForm.Hide
     Documents.Open filename:=getFileName(name)
End Sub

Function getFileName(rawName As String)
 getFileName = Replace(rawName, "[", "")
 getFileName = Replace(getFileName, "]", "")
 getFileName = Replace(getFileName, " ", "")
 getFileName = "C:\XRAY\forms\" & getFileName & ".dotm"
 
End Function
Function addNewDocument(name As String)
newDocument = Documents.Add(Template:="Normal", NewTemplate:=False, DocumentType:=0)
 ActiveDocument.SaveAs2 filename:=name, FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=True, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, CompatibilityMode:=15
 ActiveDocument.Close
End Function
Sub createDocumentFromButtonName(CommandButtonName As String)
    Dim name As String
    name = getFileName(CommandButtonName)
    addNewDocument (name)
End Sub


