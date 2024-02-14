VERSION 5.00
Begin VB.Form frmKVI 
   BackColor       =   &H00423532&
   BorderStyle     =   0  'None
   Caption         =   "kvi"
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8265
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00423532&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   3270
      Left            =   -15
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   8295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00423532&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   645
      Left            =   90
      TabIndex        =   0
      Top             =   225
      Width           =   8115
   End
End
Attribute VB_Name = "frmKVI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
    Me.KeyPreview = True
    Fill_List listCommands
End Sub

Private Sub Fill_List(dict As Dictionary)
    List1.Clear
    Dim currKey As Variant
    For Each currKey In dict.Keys
        If Left(currKey, 9) <> "Separator" Then List1.AddItem currKey
    Next currKey

End Sub
Private Sub Text1_Change()
    Dim dict As Dictionary
    If Len(Trim(Text1.Text)) > 0 Then
        Set dict = FilterDictionaryByKey(listCommands, Text1.Text)
        Fill_List dict
        Set dict = Nothing
    Else
        Fill_List listCommands
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If List1.ListCount = 1 Then
            ExecCommand listCommands.Item(List1.List(0))
        End If
    End If
End Sub
