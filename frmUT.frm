VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmUT 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistem Informasi Index Prestasi"
   ClientHeight    =   6030
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   7335
   Icon            =   "frmUT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5355
      ScaleWidth      =   7035
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.TextBox txtIsi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   4
         Left            =   5880
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   840
         Width           =   600
      End
      Begin VB.TextBox txtPros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2640
         TabIndex        =   11
         Text            =   "0"
         Top             =   0
         Width           =   630
      End
      Begin VB.TextBox txtPros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   3720
         TabIndex        =   10
         Text            =   "0"
         Top             =   0
         Width           =   630
      End
      Begin VB.TextBox txtPros 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   4800
         TabIndex        =   9
         Text            =   "0"
         Top             =   0
         Width           =   630
      End
      Begin VB.TextBox txtIsi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   840
         Width           =   2400
      End
      Begin VB.TextBox txtIsi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   2520
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox txtIsi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   2
         Left            =   3600
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   840
         Width           =   840
      End
      Begin VB.TextBox txtIsi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   3
         Left            =   4680
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Width           =   1080
      End
      Begin VB.CommandButton cmdIsi 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6600
         TabIndex        =   4
         Top             =   840
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Hapus Semua"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   2
         Top             =   4800
         Width           =   1935
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Hitung IP"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   4560
         TabIndex        =   1
         Top             =   4800
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid msf 
         Height          =   3375
         Left            =   0
         TabIndex        =   3
         Top             =   1320
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   5953
         _Version        =   393216
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "SKS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Prosentase:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   19
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "UTS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   2640
         TabIndex        =   18
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "UAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3720
         TabIndex        =   17
         Top             =   480
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "TUGAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   4680
         TabIndex        =   16
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "MATA KULIAH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   3360
         TabIndex        =   14
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   4440
         TabIndex        =   13
         Top             =   0
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   5520
         TabIndex        =   12
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright (c) Cantika Software 2014"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      MouseIcon       =   "frmUT.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   5640
      Width           =   4455
   End
   Begin VB.Menu mnuP 
      Caption         =   "Atur Kategori"
   End
End
Attribute VB_Name = "frmUT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdIsi_Click()
    If txtIsi(0).Text <> "" And txtIsi(1).Text <> "" And txtIsi(2).Text <> "" And txtIsi(3).Text <> "" And txtIsi(4).Text <> "" Then
        msf.Rows = msf.Rows + 1
        msf.TextMatrix(msf.Rows - 2, 0) = msf.Rows - 2
        msf.TextMatrix(msf.Rows - 2, 1) = txtIsi(0).Text
        msf.TextMatrix(msf.Rows - 2, 2) = txtIsi(1).Text
        msf.TextMatrix(msf.Rows - 2, 3) = txtIsi(2).Text
        msf.TextMatrix(msf.Rows - 2, 4) = txtIsi(3).Text
        msf.TextMatrix(msf.Rows - 2, 8) = txtIsi(4).Text
        msf.TextMatrix(msf.Rows - 2, 5) = (Val(txtIsi(1).Text) * Val(txtPros(1).Text) / 100) + _
        (Val(txtIsi(2).Text) * Val(txtPros(2).Text) / 100) + (Val(txtIsi(3).Text) * Val(txtPros(3).Text) / 100)
            Dim x As Long
            Dim r As String
            x = 0
            r = GetSetting("SIIP", "Kategori", "BB" & x, "")
            Do While Not r = ""
                x = x + 1
                r = GetSetting("SIIP", "Kategori", "BB" & x, "")
            Loop
            x = x - 1
            
            For i = x To 0 Step -1
                If Val(GetSetting("SIIP", "Kategori", "BB" & i, "")) <= Val(msf.TextMatrix(msf.Rows - 2, 5)) And Val(msf.TextMatrix(msf.Rows - 2, 5)) <= Val(GetSetting("SIIP", "Kategori", "BA" & i, "")) Then
                    msf.TextMatrix(msf.Rows - 2, 6) = GetSetting("SIIP", "Kategori", "H" & i, "")
                    msf.TextMatrix(msf.Rows - 2, 7) = GetSetting("SIIP", "Kategori", "NB" & i, "")
                    Exit For
                End If
            Next i
            txtIsi(0).Text = ""
            txtIsi(1).Text = ""
            txtIsi(2).Text = ""
            txtIsi(3).Text = ""
            txtIsi(4).Text = ""
            txtIsi(0).SetFocus
    End If
End Sub

Private Sub Command1_Click()
    txtIsi(0).Text = ""
    txtIsi(1).Text = ""
    txtIsi(2).Text = ""
    txtIsi(3).Text = ""
    
    AturTabel
    txtPros(1).SetFocus
End Sub

Private Sub Command2_Click()
    Dim nBobot, nSks As Long
    nBobot = 0
    nSks = 0
    For i = 1 To msf.Rows - 2
        nBobot = nBobot + ((msf.TextMatrix(i, 7)) * msf.TextMatrix(i, 8))
        nSks = nSks + Val(msf.TextMatrix(i, 8))
    Next i
    MsgBox "Indeks Prestasi Anda: " & nBobot / nSks
End Sub

Private Sub Form_Load()
        txtPros(1).Text = 30
        txtPros(2).Text = 30
        txtPros(3).Text = 40
    
    For i = txtIsi.LBound To txtIsi.UBound
        txtIsi(i).Text = ""
    Next i
    AturTabel
    CekSeting
End Sub

Sub CekSeting()
    If GetSetting("SIIP", "Kategori", "BB0", "") = "" Then
        SaveSetting "SIIP", "Kategori", "BB0", "80"
        SaveSetting "SIIP", "Kategori", "BA0", "100"
        SaveSetting "SIIP", "Kategori", "H0", "A"
        SaveSetting "SIIP", "Kategori", "NB0", "4,0"
        
        SaveSetting "SIIP", "Kategori", "BB1", "75"
        SaveSetting "SIIP", "Kategori", "BA1", "79"
        SaveSetting "SIIP", "Kategori", "H1", "B+"
        SaveSetting "SIIP", "Kategori", "NB1", "3,5"
        
        SaveSetting "SIIP", "Kategori", "BB2", "65"
        SaveSetting "SIIP", "Kategori", "BA2", "74"
        SaveSetting "SIIP", "Kategori", "H2", "B"
        SaveSetting "SIIP", "Kategori", "NB2", "3,0"
        
        SaveSetting "SIIP", "Kategori", "BB3", "60"
        SaveSetting "SIIP", "Kategori", "BA3", "64"
        SaveSetting "SIIP", "Kategori", "H3", "C+"
        SaveSetting "SIIP", "Kategori", "NB3", "2,5"
        
        SaveSetting "SIIP", "Kategori", "BB4", "55"
        SaveSetting "SIIP", "Kategori", "BA4", "59"
        SaveSetting "SIIP", "Kategori", "H4", "C"
        SaveSetting "SIIP", "Kategori", "NB4", "2,0"

        SaveSetting "SIIP", "Kategori", "BB5", "40"
        SaveSetting "SIIP", "Kategori", "BA5", "54"
        SaveSetting "SIIP", "Kategori", "H5", "D"
        SaveSetting "SIIP", "Kategori", "NB5", "1,0"

        SaveSetting "SIIP", "Kategori", "BB6", "0"
        SaveSetting "SIIP", "Kategori", "BA6", "39"
        SaveSetting "SIIP", "Kategori", "H6", "E"
        SaveSetting "SIIP", "Kategori", "NB6", "0"
    End If
End Sub

Sub AturTabel()
    msf.Rows = 2
    msf.Cols = txtIsi.Count + 4
    msf.ColWidth(0) = 450
    msf.ColWidth(1) = 2440
    msf.ColWidth(2) = 550
    msf.ColWidth(3) = 550
    msf.ColWidth(4) = 650
    msf.ColWidth(5) = 550
    msf.ColWidth(6) = 550
    msf.ColWidth(7) = 550
    msf.ColWidth(8) = 550
    
    msf.TextMatrix(0, 0) = "No"
    msf.TextMatrix(0, 1) = "Mata Kuliah"
    msf.TextMatrix(0, 2) = "UTS"
    msf.TextMatrix(0, 3) = "UAS"
    msf.TextMatrix(0, 4) = "TUGAS"
    msf.TextMatrix(0, 5) = "NA"
    msf.TextMatrix(0, 6) = "KAT"
    msf.TextMatrix(0, 7) = "NB"
    msf.TextMatrix(0, 8) = "SKS"

    msf.TextMatrix(1, 0) = ""
    msf.TextMatrix(1, 1) = ""
    msf.TextMatrix(1, 2) = ""
    msf.TextMatrix(1, 3) = ""
    msf.TextMatrix(1, 4) = ""
    msf.TextMatrix(1, 5) = ""
    msf.TextMatrix(1, 6) = ""
    msf.TextMatrix(1, 7) = ""
    msf.TextMatrix(1, 8) = ""
End Sub

Private Sub Label2_Click()
    Load frmAbout
    frmAbout.Show 1, Me
End Sub

Private Sub mnuP_Click()
    Load frmP
    frmP.Show
End Sub

Private Sub txtIsi_KeyPress(Index As Integer, KeyAscii As Integer)
If (Index <> 0) Then
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8) Then
    KeyAscii = 0
    End If
End If
End Sub

Private Sub txtPros_KeyPress(Index As Integer, KeyAscii As Integer)
    If Not ((KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Or KeyAscii = 8) Then
    KeyAscii = 0
    End If
End Sub
