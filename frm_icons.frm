VERSION 5.00
Begin VB.Form frm_icons 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Available Icons"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_help 
      Caption         =   "Help me!"
      Height          =   1455
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   5655
      Begin VB.Label lbl_helpme 
         Alignment       =   2  'Center
         Caption         =   $"frm_icons.frx":0000
         Height          =   1215
         Left            =   1140
         TabIndex        =   2
         Top             =   180
         Width           =   4395
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frm_icons.frx":0144
         Top             =   540
         Width           =   480
      End
   End
   Begin VB.TextBox txt_avicons 
      Height          =   4695
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1560
      Width           =   5655
   End
End
Attribute VB_Name = "frm_icons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    load_settings
    read_avicons
End Sub

Private Sub read_avicons()
    For av_icons = 0 To noof_icons - 1
        txt_avicons.Text = txt_avicons.Text & msg_icons(av_icons).icon_recogstr & "     Gets replaced by a " & msg_icons(av_icons).icon_description & vbCrLf
    Next av_icons
End Sub

Private Sub load_settings()
    With frm_icons
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
End Sub
