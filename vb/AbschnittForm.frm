VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AbschnittForm 
   Caption         =   "Extra Abschnitt"
   ClientHeight    =   2610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   OleObjectBlob   =   "AbschnittForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AbschnittForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private abschnittPresenterObject As abschnittPresenter

Private Sub ComboBox1_Change()
    Dim items() As String
    items = abschnittPresenterObject.getFile(ComboBox3.Text, ComboBox1.Text)
    abschnittPresenterObject.setItems ComboBox2, items

End Sub

Private Sub ComboBox3_Change()
    Dim items() As String
    items = abschnittPresenterObject.getSubFolderNames(ComboBox3.Text)
    abschnittPresenterObject.setItems ComboBox1, items
End Sub

Private Sub CommandButton1_Click()
    Dim succeed As Boolean
    succeed = abschnittPresenterObject.insertSection(ComboBox3.Text, ComboBox1.Text, ComboBox2.Text)
    If (succeed) Then
        AbschnittForm.Hide
    End If
End Sub

Private Sub UserForm_Click()

End Sub



Private Sub UserForm_Initialize()

    Set abschnittPresenterObject = New abschnittPresenter
    Dim items() As String
     items = abschnittPresenterObject.getSubFolderNames("")
     abschnittPresenterObject.setItems ComboBox3, items

End Sub

