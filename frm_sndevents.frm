VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_sndevents 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Sound Events"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5730
   Icon            =   "frm_sndevents.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frm_help 
      Caption         =   "Help me!"
      Height          =   1155
      Left            =   60
      TabIndex        =   6
      Top             =   0
      Width           =   5595
      Begin VB.Image Image1 
         Height          =   480
         Left            =   240
         Picture         =   "frm_sndevents.frx":0CCA
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lbl_helpme 
         Alignment       =   2  'Center
         Caption         =   $"frm_sndevents.frx":110C
         Height          =   855
         Left            =   1080
         TabIndex        =   7
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.Frame fra_soundoptions 
      Caption         =   "Sound Options"
      Height          =   915
      Left            =   60
      TabIndex        =   3
      Top             =   4920
      Width           =   2535
      Begin VB.CommandButton cmd_test 
         Caption         =   "Test"
         Height          =   555
         Left            =   1320
         Picture         =   "frm_sndevents.frx":11E4
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_change 
         Caption         =   "Change"
         Height          =   555
         Left            =   120
         Picture         =   "frm_sndevents.frx":12E6
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   3360
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton cmd_ok 
      Caption         =   "Ok"
      Height          =   555
      Left            =   4560
      TabIndex        =   1
      Top             =   5160
      Width           =   1095
   End
   Begin MSComctlLib.TreeView trv_sndevents 
      Height          =   3675
      Left            =   60
      TabIndex        =   0
      Top             =   1200
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   6482
      _Version        =   393217
      Style           =   7
      Checkboxes      =   -1  'True
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   60
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frm_sndevents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancel_Click()
    Unload Me
End Sub

Private Sub cmd_change_Click()
    With trv_sndevents
        If .Nodes.Count > 0 Then
            If .SelectedItem.index <> 0 Then
                change_sound .SelectedItem.index
            End If
        End If
    End With
End Sub



Private Sub cmd_ok_Click()
    For saveevents = 0 To noof_events - 1
        snd_events(saveevents).snd_enabled = trv_sndevents.Nodes.Item(saveevents + 1).Checked
        snd_events(saveevents).save App.ProductName
    Next saveevents
    Unload Me
End Sub

Private Sub cmd_test_Click()
    With trv_sndevents
        If .Nodes.Count > 0 Then
            If .SelectedItem.index <> 0 Then
                snd_events(.SelectedItem.index - 1).start
            End If
        End If
    End With
End Sub

Private Sub Form_Load()
    load_settings
    refresh_events
End Sub

Private Sub load_settings()
    With frm_sndevents
        load_window (.Caption)
        If win_top <> 0 Then .Top = win_top
        If win_left <> 0 Then .Left = win_left
    End With
End Sub

Private Sub save_settings()
    save_window Me.Caption, Me.Top, Me.Left
End Sub

Private Sub refresh_events()
    trv_sndevents.Nodes.Clear
    For addnodes = 1 To noof_events
        With trv_sndevents
            .Nodes.Add , , snd_events(addnodes - 1).snd_name, snd_events(addnodes - 1).snd_name
            .Nodes.Item(addnodes).Checked = snd_events(addnodes - 1).snd_enabled
        End With
    Next addnodes
End Sub

Private Sub change_sound(index As Integer)
    Dim filename As String
        With cdlg
            .CancelError = True
            On Error GoTo ErrHandler
            .Flags = cdlOFNHideReadOnly
            .Filter = "Wave File (*.wav)|*.wav"
            .ShowSave
            snd_events(index - 1).filename .filename
        End With
    Exit Sub
ErrHandler:
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    save_settings
End Sub

Private Sub trv_sndevents_NodeCheck(ByVal Node As MSComctlLib.Node)
    With trv_sndevents
        If .SelectedItem.index > 0 Then snd_events(.SelectedItem.index).snd_enabled = .SelectedItem.Checked
    End With
End Sub
