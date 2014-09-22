VERSION 5.00
Begin VB.Form frmP 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Pengaturan Kategori Nilai"
   ClientHeight    =   1500
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1500
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdK 
      Caption         =   "Kurangi Kategori"
      Height          =   315
      Left            =   2040
      TabIndex        =   10
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdS 
      Caption         =   "Simpan"
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   1200
      Width           =   4095
   End
   Begin VB.CommandButton cmdT 
      Caption         =   "Tambah Kategori"
      Height          =   315
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   2055
   End
   Begin VB.TextBox txtNB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2880
      TabIndex        =   6
      Top             =   720
      Width           =   1110
   End
   Begin VB.TextBox txtH 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   630
   End
   Begin VB.TextBox txtBA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   720
      Width           =   390
   End
   Begin VB.TextBox txtBB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   390
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   48.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   600
      TabIndex        =   8
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "RANGE"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "NILAI BOBOT"
      Height          =   255
      Left            =   2880
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "HURUF"
      Height          =   255
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   735
   End
End
Attribute VB_Name = "frmP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private tIndex As Long

Sub TambahKategori()
    tIndex = txtBB.UBound
    If txtBB(tIndex).Text = "" And txtBA(tIndex).Text = "" _
    And txtH(tIndex).Text = "" And txtNB(tIndex).Text = "" Then
        Exit Sub
    Else
        Load txtBB(tIndex + 1)
        Load txtBA(tIndex + 1)
        Load txtH(tIndex + 1)
        Load txtNB(tIndex + 1)
        Load lbl(tIndex + 1)
        
        txtBB(tIndex + 1).Visible = True
        txtBA(tIndex + 1).Visible = True
        txtH(tIndex + 1).Visible = True
        txtNB(tIndex + 1).Visible = True
        lbl(tIndex + 1).Visible = True

        txtBB(tIndex + 1).Top = txtBB(tIndex).Top + txtBB(tIndex).Height
        txtBA(tIndex + 1).Top = txtBA(tIndex).Top + txtBA(tIndex).Height
        txtH(tIndex + 1).Top = txtH(tIndex).Top + txtH(tIndex).Height
        txtNB(tIndex + 1).Top = txtNB(tIndex).Top + txtNB(tIndex).Height
        Dim n As Long
        n = txtNB(tIndex).Top - lbl(tIndex).Top
        lbl(tIndex + 1).Top = txtNB(tIndex + 1).Top - n
        
        txtBB(tIndex + 1).Text = ""
        txtBA(tIndex + 1).Text = ""
        txtH(tIndex + 1).Text = ""
        txtNB(tIndex + 1).Text = ""
        
        cmdS.Top = cmdS.Top + txtNB(tIndex).Height
        Me.Height = Me.Height + txtNB(tIndex).Height
    End If
End Sub

Sub KurangKategori()
    tIndex = txtBB.UBound
    If tIndex <> 0 Then
        Unload txtBB(tIndex)
        Unload txtBA(tIndex)
        Unload txtH(tIndex)
        Unload txtNB(tIndex)
        Unload lbl(tIndex)
                
        cmdS.Top = cmdS.Top - txtNB(tIndex - 1).Height
        Me.Height = Me.Height - txtNB(tIndex - 1).Height
    End If
End Sub

Sub Simpan()
    Dim x As Long
    Dim r As String
    x = 0
    r = GetSetting("SIIP", "Kategori", "BB" & x, "")
    Do While Not r = ""
        DeleteSetting "SIIP", "Kategori", "BB" & x
        DeleteSetting "SIIP", "Kategori", "BA" & x
        DeleteSetting "SIIP", "Kategori", "H" & x
        DeleteSetting "SIIP", "Kategori", "NB" & x
        x = x + 1
        r = GetSetting("SIIP", "Kategori", "BB" & x, "")
    Loop

    For i = txtBA.LBound To txtBB.UBound
        If txtBB(i).Text <> "" And txtBA(i).Text <> "" _
        And txtH(i).Text <> "" And txtNB(i).Text <> "" Then
            SaveSetting "SIIP", "Kategori", "BB" & i, txtBB(i).Text
            SaveSetting "SIIP", "Kategori", "BA" & i, txtBA(i).Text
            SaveSetting "SIIP", "Kategori", "H" & i, txtH(i).Text
            SaveSetting "SIIP", "Kategori", "NB" & i, txtNB(i).Text
        End If
    Next i
End Sub

Private Sub cmdK_Click()
    KurangKategori
End Sub

Private Sub cmdS_Click()
    Simpan
End Sub

Private Sub cmdT_Click()
    TambahKategori
End Sub

Private Sub Form_Load()
    Dim x As Long
    Dim r As String
    x = 0
    r = GetSetting("SIIP", "Kategori", "BB" & x, "")
    Do While Not r = ""
        txtBB(x).Text = GetSetting("SIIP", "Kategori", "BB" & x, "")
        txtBA(x).Text = GetSetting("SIIP", "Kategori", "BA" & x, "")
        txtH(x).Text = GetSetting("SIIP", "Kategori", "H" & x, "")
        txtNB(x).Text = GetSetting("SIIP", "Kategori", "NB" & x, "")
        TambahKategori
        x = x + 1
        r = GetSetting("SIIP", "Kategori", "BB" & x, "")
    Loop
    KurangKategori
End Sub
