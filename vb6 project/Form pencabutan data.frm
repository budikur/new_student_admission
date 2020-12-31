VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Form2 
   Caption         =   "PENCABUTAN DATA"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8130
   LinkTopic       =   "Form2"
   ScaleHeight     =   4260
   ScaleWidth      =   8130
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox txt_noDaftar 
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.CommandButton cmd_cabut 
      Caption         =   "PENCABUTAN DATA"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Width           =   2775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   8535
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   15055
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Caption         =   "Nomor Pendaftaran"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "PENCARIAN ATAU PENCABUTAN DATA CALON PESERTA DIDIK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   15015
   End
   Begin VB.Menu ENTRY_DATA 
      Caption         =   "&ENTRY DATA"
   End
   Begin VB.Menu CABUT_DATA 
      Caption         =   "&CARI/CABUT DATA"
   End
   Begin VB.Menu PROSES_DATA 
      Caption         =   "&PROSES DATA"
   End
   Begin VB.Menu KELUAR 
      Caption         =   "&KELUAR"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ENTRY_DATA_Click()
    Form2.Hide
    Form1.Form_Load
End Sub

Private Sub KELUAR_Click()
    End
End Sub

Private Sub PROSES_DATA_Click()
    Form2.Hide
    Form3.Show
End Sub
