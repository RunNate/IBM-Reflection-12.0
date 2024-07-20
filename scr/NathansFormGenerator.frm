VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NathansFormGenerator 
   Caption         =   "Form Generator V1.0"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4800
   OleObjectBlob   =   "NathansFormGenerator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NathansFormGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    ComboBox1.AddItem "Taxi Travel Form"
    ComboBox1.AddItem "ADL Request Form"
End Sub

Private Sub Cancel_Click()
    Unload Me
End Sub

Private Sub Generate_Click()
    Dim selectedFunction As String
    selectedFunction = ComboBox1.Value
    
    Select Case selectedFunction
        Case "Taxi Travel Form"
            customMacros.createTaxiRequestForm
            Unload Me
        Case "ADL Request Form"
            customMacros.createADLForm
            Unload Me
        Case Else
            MsgBox "Please select a letter from the list."
    End Select
End Sub


