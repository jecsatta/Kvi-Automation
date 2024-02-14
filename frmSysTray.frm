VERSION 5.00
Begin VB.Form frmSysTray 
   Appearance      =   0  'Flat
   Caption         =   "Hard Drive"
   ClientHeight    =   690
   ClientLeft      =   1425
   ClientTop       =   2295
   ClientWidth     =   1575
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   46
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   105
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   840
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   0
      Top             =   120
      Width           =   480
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   120
   End
   Begin VB.Menu mnuSysTray 
      Caption         =   "SysTray"
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Close Menu"
         Index           =   998
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Exit"
         Index           =   999
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents SysTray As clsSysTray
Attribute SysTray.VB_VarHelpID = -1
Private Sub LoadMenu()
    Dim elem As Variant
    Dim i As Integer
    i = 0
    For Each elem In listCommands.Keys
       Load mnuPopup(i)
       If Left(elem, 9) = "Separator" Then
        mnuPopup(i).Caption = "-"
       Else
        mnuPopup(i).Caption = "" & elem
       End If
       i = i + 1
    Next
End Sub
Private Sub Form_Load()
    Set SysTray = New clsSysTray
    Me.WindowState = vbMinimized
    
    DoEvents
   
    mnuPopup(999).Caption = App_ExitCaption
    mnuPopup(998).Caption = App_CloseMenuCaption
    LoadMenu
    Me.Hide
    
    SysTray.Init Me, App_Hint
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set SysTray = Nothing
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
    SysTray.MouseMove Button, X, Me
End Sub

Private Sub mnuPopup_Click(Index As Integer)
    Dim elem    As Variant
    Dim vPrefix As String
    
    If Me.mnuPopup(Index).Caption = App_ExitCaption Then
        Unload frmKVI
        Unload Me
    Else
        If Not listCommands.Exists(Me.mnuPopup(Index).Caption) Then Exit Sub
        vPrefix = "" & listCommands.Item(Me.mnuPopup(Index).Caption)
        For Each elem In listCommandsItems.Keys
            If Left("" & elem, Len(vPrefix) + 1) = vPrefix & "_" Then
                Shell listCommandsItems.Item(elem), vbNormalFocus
            End If
        Next
    End If
End Sub


Private Sub SysTray_DoubleClick()
    Dim elem As Variant
    Dim i As Integer
    i = 0
    For Each elem In listCommandsDblClick.Items
        Shell "" & elem
    Next
End Sub

Private Sub SysTray_LeftClick()
    frmKVI.Show
End Sub

Private Sub SysTray_RightClick()
    PopupMenu Me.mnuSysTray
End Sub

