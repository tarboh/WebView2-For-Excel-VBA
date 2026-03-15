VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9120.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12420
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit



Private Sub CommandButton1_Click()
    
    Dim url As String
    url = TextBox1.Text
    
    Call Module1.WV2.Navigate(url)

End Sub


Private Sub CommandButton2_Click()
    Dim script As String
    script = InputBox("Input JavaScript")
    
    Call Module1.WV2.ExecuteScript(script)
    
End Sub

Private Sub UserForm_Activate()
    Call Module1.WebView2錬成
End Sub

