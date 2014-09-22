VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tentang Program"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5535
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   2640
      Top             =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "MENERIMA PESANAN APLIKASI"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   1800
      Width           =   5535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "APLIKASI INI OPEN SOURCE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label3 
      Caption         =   "Kontak: 085731292834; 03133225290; PIN: 326E7155"
      Height          =   255
      Left            =   1440
      TabIndex        =   2
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Alamat: Jetis Kulon 1/63 Surabaya"
      Height          =   255
      Left            =   1440
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "Programmer: Muhammad Hassanuddin"
      Height          =   255
      Left            =   1440
      TabIndex        =   0
      Top             =   0
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   1080
      Left            =   0
      Picture         =   "frmAbout.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1320
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
    If Label5.Visible = True Then
        Label5.Visible = False
    Else
        Label5.Visible = True
    End If
End Sub
