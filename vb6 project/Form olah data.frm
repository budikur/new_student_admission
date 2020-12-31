VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "DISPLAY HASIL PENGOLAHAN DATA"
   ClientHeight    =   5535
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   8475
   LinkTopic       =   "Form3"
   ScaleHeight     =   5535
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmd_proses_kedua 
      Caption         =   "PROSES KE 2"
      Height          =   495
      Left            =   5160
      TabIndex        =   36
      Top             =   600
      Width           =   2415
   End
   Begin VB.CommandButton cmd_proses_lanjutan 
      Caption         =   "PROSES LOOPING"
      Enabled         =   0   'False
      Height          =   495
      Left            =   7680
      TabIndex        =   35
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox T_ggl_AK 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14520
      TabIndex        =   33
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox T_trima_AK 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14520
      TabIndex        =   31
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox T_ggl_TKJ 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14520
      TabIndex        =   29
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox T_trima_TKJ 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14520
      TabIndex        =   27
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox T_ggl_PJ 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   25
      Top             =   7920
      Width           =   615
   End
   Begin VB.TextBox T_ggl_AP 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   23
      Top             =   5880
      Width           =   615
   End
   Begin VB.TextBox T_trima_PJ 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   21
      Top             =   3240
      Width           =   615
   End
   Begin VB.TextBox T_trima_AP 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   19
      Top             =   1200
      Width           =   615
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton C_ggl_tkj 
      Caption         =   "TEKNIK KOMPUTER DAN JARINGAN"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   7920
      Width           =   5055
   End
   Begin VB.CommandButton C_ggl_ak 
      Caption         =   "AKUNTANSI"
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   5880
      Width           =   5055
   End
   Begin VB.CommandButton C_ggl_pj 
      Caption         =   "PENJUALAN/MARKETING"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   7920
      Width           =   5055
   End
   Begin VB.CommandButton C_ggl_ap 
      Caption         =   "ADMINISTRASI PERKANTORAN"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5880
      Width           =   5055
   End
   Begin VB.CommandButton C_trm_tkj 
      Caption         =   "TEKNIK KOMPUTER DAN JARINGAN DENGAN JUMLAH PAGU 38"
      Height          =   375
      Left            =   7680
      TabIndex        =   5
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton C_trm_ak 
      Caption         =   "AKUNTANSI DENGAN JUMLAH PAGU 41"
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   1200
      Width           =   5055
   End
   Begin VB.CommandButton C_trm_pj 
      Caption         =   "PENJUALAN/MARKETING DENGAN JUMLAH PAGU 56"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   5055
   End
   Begin VB.CommandButton C_trm_ap 
      Caption         =   "ADMINISTRASI PERKANTORAN DENGAN JUMLAH PAGU 38"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5055
   End
   Begin MSDataGridLib.DataGrid DataGrid2 
      Height          =   1695
      Left            =   7680
      TabIndex        =   12
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid3 
      Height          =   1695
      Left            =   120
      TabIndex        =   13
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid4 
      Height          =   1695
      Left            =   7680
      TabIndex        =   14
      Top             =   3600
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid5 
      Height          =   1695
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid6 
      Height          =   1695
      Left            =   7680
      TabIndex        =   16
      Top             =   6240
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid7 
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   8280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSDataGridLib.DataGrid DataGrid8 
      Height          =   1695
      Left            =   7680
      TabIndex        =   18
      Top             =   8280
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2990
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1057
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label8 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   13560
      TabIndex        =   34
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label7 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   13560
      TabIndex        =   32
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   13560
      TabIndex        =   30
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   13560
      TabIndex        =   28
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   6000
      TabIndex        =   24
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   6000
      TabIndex        =   22
      Top             =   3360
      Width           =   975
   End
   Begin VB.Label Label27 
      Caption         =   "Jumlah data"
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "DAFTAR CPD YANG BELUM DITERIMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5520
      Width           =   5295
   End
   Begin VB.Label Label14 
      Caption         =   "DAFTAR CPD YANG DITERIMA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "DAFTAR SELURUH DATA HASIL PEMROSESAN"
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
      TabIndex        =   0
      Top             =   120
      Width           =   15015
   End
   Begin VB.Menu ENTRY_DATA 
      Caption         =   "&ENTRY DATA"
   End
   Begin VB.Menu PROSES_DATA 
      Caption         =   "&PROSES DATA"
   End
   Begin VB.Menu KELUAR 
      Caption         =   "&KELUAR"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_trm_ap As New ADODB.Recordset
Public rs_trm_ak As New ADODB.Recordset
Public rs_trm_pj As New ADODB.Recordset
Public rs_trm_tkj As New ADODB.Recordset
Public rs_ggl_ap As New ADODB.Recordset
Public rs_ggl_ak As New ADODB.Recordset
Public rs_ggl_pj As New ADODB.Recordset
Public rs_ggl_tkj As New ADODB.Recordset


Private Sub cmd_proses_kedua_Click()
'========================================================================DATA DITRIMA SEMENTARA
    '=============================================
    'MENUTUP RECORDSET UNTUK QUERY YANG LAMA
    '=============================================
    rs_trm_ap.Close
    rs_trm_ak.Close
    rs_trm_pj.Close
    rs_trm_tkj.Close
    '=============================================
    'MEMBUKA RECORDSET UNTUK QUERY YANG BARU
    '=============================================
    rs_trm_ap.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.jurusan = 'AP'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_ap.Requery
    rs_trm_ak.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.jurusan = 'AK'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_ak.Requery
    rs_trm_pj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.jurusan = 'PJ'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_pj.Requery
    rs_trm_tkj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.jurusan = 'TKJ'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_tkj.Requery
    '=============================================
    'PENGURUTAN BERDASARKAN TOTAL BOBOT
    '=============================================
    rs_trm_ap.Sort = "score_nilai desc"
    rs_trm_ak.Sort = "score_nilai desc"
    rs_trm_pj.Sort = "score_nilai desc"
    rs_trm_tkj.Sort = "score_nilai desc"
    '=============================================
    'PEMBERIAN RANKING
    '=============================================
    On Error Resume Next
    rs_trm_ap.MoveFirst
    For i = 1 To rs_trm_ap.RecordCount
        rs_trm_ap(0) = i
        rs_trm_ap.Update
        rs_trm_ap.MoveNext
    Next i
    rs_trm_ak.MoveFirst
    For i = 1 To rs_trm_ak.RecordCount
        rs_trm_ak(0) = i
        rs_trm_ak.Update
        rs_trm_ak.MoveNext
    Next i
    On Error Resume Next
    rs_trm_pj.MoveFirst
    For i = 1 To rs_trm_pj.RecordCount
        rs_trm_pj(0) = i
        rs_trm_pj.Update
        rs_trm_pj.MoveNext
    Next i
    rs_trm_tkj.MoveFirst
    For i = 1 To rs_trm_tkj.RecordCount
        rs_trm_tkj(0) = i
        rs_trm_tkj.Update
        rs_trm_tkj.MoveNext
    Next i
    
    
    '=============================================
    'POSISI KURSOR DITARUH BAWAH AGAR DATA TERBAWAH KELIHATAN
    '=============================================
    rs_trm_ap.MoveLast
    rs_trm_ak.MoveLast
    rs_trm_pj.MoveLast
    rs_trm_tkj.MoveLast
    '=============================================
    'MENGHITUNG JUMLAH DATA
    '=============================================
    T_trima_AP.Text = rs_trm_ap.RecordCount
    T_trima_AK.Text = rs_trm_ak.RecordCount
    T_trima_PJ.Text = rs_trm_pj.RecordCount
    T_trima_TKJ.Text = rs_trm_tkj.RecordCount
'========================================================================DATA GAGAL
    '=============================================
    'MENUTUP RECORDSET UNTUK QUERY GAGAL YANG LAMA
    '=============================================
    If rs_ggl_ap.State = 1 Then
        rs_ggl_ap.Close
    End If
    If rs_ggl_ak.State = 1 Then
        rs_ggl_ak.Close
    End If
    If rs_ggl_pj.State Then
        rs_ggl_pj.Close
    End If
    If rs_ggl_tkj.State Then
        rs_ggl_tkj.Close
    End If
    '====================================================================
    'MENYIMPAN RECORD DI MASING2 VARIABEL (trm=diterima, ggl=gagal)
    'DG MENGELOMPOKKAN DATA BERDASARKAN PILIHAN I
    '====================================================================
    rs_ggl_ap.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_2 = 'AP' and table1.jurusan='BELUM DITERIMA'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_ggl_ak.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_2 = 'AK' and table1.jurusan='BELUM DITERIMA'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_ggl_pj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_2 = 'PJ' and table1.jurusan='BELUM DITERIMA'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_ggl_tkj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_2 = 'TKJ' and table1.jurusan='BELUM DITERIMA'", Form1.con, adOpenKeyset, adLockOptimistic
    Set DataGrid5.DataSource = rs_ggl_ap
    Set DataGrid6.DataSource = rs_ggl_ak
    Set DataGrid7.DataSource = rs_ggl_pj
    Set DataGrid8.DataSource = rs_ggl_tkj
    T_ggl_AP.Text = rs_ggl_ap.RecordCount
    T_ggl_AK.Text = rs_ggl_ak.RecordCount
    T_ggl_PJ.Text = rs_ggl_pj.RecordCount
    T_ggl_TKJ.Text = rs_ggl_tkj.RecordCount
    '=============================================
    'PENGURUTAN BERDASARKAN TOTAL BOBOT
    '=============================================
    rs_ggl_ap.Sort = "score_nilai desc"
    rs_ggl_ak.Sort = "score_nilai desc"
    rs_ggl_pj.Sort = "score_nilai desc"
    rs_ggl_tkj.Sort = "score_nilai desc"
    '=============================================
    'PEMBERIAN RANKING
    '=============================================
    On Error Resume Next
    rs_ggl_ap.MoveFirst
    For i = 1 To rs_ggl_ap.RecordCount
        rs_ggl_ap(0) = i
        rs_ggl_ap.Update
        rs_ggl_ap.MoveNext
    Next i
    rs_ggl_ak.MoveFirst
    For i = 1 To rs_ggl_ak.RecordCount
        rs_ggl_ak(0) = i
        rs_ggl_ak.Update
        rs_ggl_ak.MoveNext
    Next i
    On Error Resume Next
    rs_ggl_pj.MoveFirst
    For i = 1 To rs_ggl_pj.RecordCount
        rs_ggl_pj(0) = i
        rs_ggl_pj.Update
        rs_ggl_pj.MoveNext
    Next i
    rs_ggl_tkj.MoveFirst
    For i = 1 To rs_ggl_tkj.RecordCount
        rs_ggl_tkj(0) = i
        rs_ggl_tkj.Update
        rs_ggl_tkj.MoveNext
    Next i
    '=============================================
    'POSISI KURSOR DITARUH ATAS AGAR DATA TERATAS KELIHATAN
    '=============================================
    rs_ggl_ap.MoveFirst
    rs_ggl_ak.MoveFirst
    rs_ggl_pj.MoveFirst
    rs_ggl_tkj.MoveFirst
    
    cmd_proses_kedua.Enabled = False
    cmd_proses_lanjutan.Enabled = True
End Sub

Public Sub cmd_proses_lanjutan_Click()
'    cmd_proses_lanjutan.Caption = "PROSES KE " & Right(cmd_proses_lanjutan.Caption, 1) + 1
    
    '=====================================================
    'SELEKSI BERDASARKAN PILIHAN 2 UNTUK JURUSAN AP
    '=====================================================
    If T_trima_AP.Text < Right(C_trm_ap.Caption, 2) Then    'JIKA JURUSAN BELUM MENCAPAI PAGU DARI PILIHAN I
        rs_ggl_ap.MoveFirst
        For i = 1 To Right(C_trm_ap.Caption, 2) - T_trima_AP.Text
            rs_ggl_ap(7) = "AP"
            rs_ggl_ap.Update
            rs_ggl_ap.MoveNext
        Next i
    Else                                                    'JIKA JURUSAN SUDAH MENCAPAI PAGU DARI PILIHAN I
        rs_trm_ap.MoveLast
        rs_ggl_ap.MoveFirst
        While (rs_trm_ap(2) < rs_ggl_ap(2))
            rs_trm_ap(7) = "BELUM DITERIMA"
            rs_ggl_ap(7) = "AP"
            rs_trm_ap.MovePrevious
            rs_ggl_ap.MoveNext
        Wend
    End If
    '=====================================================
    'SELEKSI BERDASARKAN PILIHAN 2 UNTUK JURUSAN AK
    '=====================================================
    If T_trima_AK.Text < Right(C_trm_ak.Caption, 2) Then    'JIKA JURUSAN BELUM MENCAPAI PAGU DARI PILIHAN I
        rs_ggl_ak.MoveFirst
        For i = 1 To Right(C_trm_ak.Caption, 2) - T_trima_AK.Text
            rs_ggl_ak(7) = "AK"
            rs_ggl_ak.Update
            rs_ggl_ak.MoveNext
        Next i
    Else                                                    'JIKA JURUSAN SUDAH MENCAPAI PAGU DARI PILIHAN I
        rs_trm_ak.MoveLast
        rs_ggl_ak.MoveFirst
        While (rs_trm_ak(2) < rs_ggl_ak(2))
            rs_trm_ak(7) = "BELUM DITERIMA"
            rs_ggl_ak(7) = "AK"
            rs_trm_ak.MovePrevious
            rs_ggl_ak.MoveNext
        Wend
    End If
    '=====================================================
    'SELEKSI BERDASARKAN PILIHAN 2 UNTUK JURUSAN PJ
    '=====================================================
    If T_trima_PJ.Text < Right(C_trm_pj.Caption, 2) Then    'JIKA JURUSAN BELUM MENCAPAI PAGU DARI PILIHAN I
        rs_ggl_pj.MoveFirst
        For i = 1 To Right(C_trm_pj.Caption, 2) - T_trima_PJ.Text
            rs_ggl_pj(7) = "PJ"
            rs_ggl_pj.Update
            rs_ggl_pj.MoveNext
        Next i
    Else                                                    'JIKA JURUSAN SUDAH MENCAPAI PAGU DARI PILIHAN I
        rs_trm_pj.MoveLast
        rs_ggl_pj.MoveFirst
        While (rs_trm_pj(2) < rs_ggl_pj(2))
            rs_trm_pj(7) = "BELUM DITERIMA"
            rs_ggl_pj(7) = "PJ"
            rs_trm_pj.MovePrevious
            rs_ggl_pj.MoveNext
        Wend
    End If
    '=====================================================
    'SELEKSI BERDASARKAN PILIHAN 2 UNTUK JURUSAN TKJ
    '=====================================================
    If T_trima_TKJ.Text < Right(C_trm_tkj.Caption, 2) Then    'JIKA JURUSAN BELUM MENCAPAI PAGU DARI PILIHAN I
        rs_ggl_tkj.MoveFirst
        For i = 1 To Right(C_trm_tkj.Caption, 2) - T_trima_TKJ.Text
            rs_ggl_tkj(7) = "TKJ"
            rs_ggl_tkj.Update
            rs_ggl_tkj.MoveNext
        Next i
    Else                                                    'JIKA JURUSAN SUDAH MENCAPAI PAGU DARI PILIHAN I
        rs_trm_tkj.MoveLast
        rs_ggl_tkj.MoveFirst
        While (rs_trm_tkj(2) < rs_ggl_tkj(2))
            rs_trm_tkj(7) = "BELUM DITERIMA"
            rs_ggl_tkj(7) = "TKJ"
            rs_trm_tkj.MovePrevious
            rs_ggl_tkj.MoveNext
        Wend
    End If
    
    cmd_proses_kedua_Click
End Sub

Private Sub ENTRY_DATA_Click()
    Form1.con.Close
    Form3.Hide
    Form1.Form_Load
    Form1.Show
    cmd_proses_kedua.Enabled = True
    cmd_proses_lanjutan.Enabled = False
End Sub

Public Sub Form_Load()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Form1.con.State = 1 Then Form1.con.Close
    Form1.con.CursorLocation = adUseClient
    Form1.con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\dsa\data calon peserta didik.mdb;"
    '====================================================================
    'MENYIMPAN RECORD DI MASING2 VARIABEL (trm=diterima, ggl=gagal)
    'DG MENGELOMPOKKAN DATA BERDASARKAN PILIHAN I
    '====================================================================
    rs_trm_ap.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_1 = 'AP'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_ak.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_1 = 'AK'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_pj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_1 = 'PJ'", Form1.con, adOpenKeyset, adLockOptimistic
    rs_trm_tkj.Open "select rangking_di_jurusan,no_daftar,score_nilai,nama_siswa,pil_1,pil_2,sign,jurusan from table1 Where table1.sign = 'TERDAFTAR' And table1.pil_1 = 'TKJ'", Form1.con, adOpenKeyset, adLockOptimistic
    Set DataGrid1.DataSource = rs_trm_ap
    Set DataGrid2.DataSource = rs_trm_ak
    Set DataGrid3.DataSource = rs_trm_pj
    Set DataGrid4.DataSource = rs_trm_tkj
    T_trima_AP.Text = rs_trm_ap.RecordCount
    T_trima_AK.Text = rs_trm_ak.RecordCount
    T_trima_PJ.Text = rs_trm_pj.RecordCount
    T_trima_TKJ.Text = rs_trm_tkj.RecordCount
    '=============================================
    'PENGURUTAN BERDASARKAN TOTAL BOBOT
    '=============================================
    rs_trm_ap.Sort = "score_nilai desc"
    rs_trm_ak.Sort = "score_nilai desc"
    rs_trm_pj.Sort = "score_nilai desc"
    rs_trm_tkj.Sort = "score_nilai desc"
    '=============================================
    'PEMBERIAN RANKING
    '=============================================
    On Error Resume Next
    rs_trm_ap.MoveFirst
    For i = 1 To rs_trm_ap.RecordCount
        rs_trm_ap(0) = i
        rs_trm_ap.Update
        rs_trm_ap.MoveNext
    Next i
    rs_trm_ak.MoveFirst
    For i = 1 To rs_trm_ak.RecordCount
        rs_trm_ak(0) = i
        rs_trm_ak.Update
        rs_trm_ak.MoveNext
    Next i
    rs_trm_pj.MoveFirst
    For i = 1 To rs_trm_pj.RecordCount
        rs_trm_pj(0) = i
        rs_trm_pj.Update
        rs_trm_pj.MoveNext
    Next i
    rs_trm_tkj.MoveFirst
    For i = 1 To rs_trm_tkj.RecordCount
        rs_trm_tkj(0) = i
        rs_trm_tkj.Update
        rs_trm_tkj.MoveNext
    Next i
    '=============================================
    'PEMBERIAN STATUS BELUM DITERIMA PADA FIELD JURUSAN UTK SEMUA DATA
    'AGAR DATA STATUS JURUSAN SBG TANDA DITERIMA VALID
    '=============================================
    rs_trm_ap.MoveFirst
    For i = 1 To rs_trm_ap.RecordCount                  'SEMUA DATA YG PIL1 AP
        rs_trm_ap(7) = "BELUM DITERIMA"
        rs_trm_ap.Update
        rs_trm_ap.MoveNext
    Next i
    rs_trm_ak.MoveFirst
    For i = 1 To rs_trm_ak.RecordCount                  'SEMUA DATA YG PIL1 AK
        rs_trm_ak(7) = "BELUM DITERIMA"
        rs_trm_ak.Update
        rs_trm_ak.MoveNext
    Next i
    rs_trm_pj.MoveFirst
    For i = 1 To rs_trm_pj.RecordCount                  'SEMUA DATA YG PIL1 PJ
        rs_trm_pj(7) = "BELUM DITERIMA"
        rs_trm_pj.Update
        rs_trm_pj.MoveNext
    Next i
    rs_trm_tkj.MoveFirst
    For i = 1 To rs_trm_tkj.RecordCount                  'SEMUA DATA YG PIL1 TKJ
        rs_trm_tkj(7) = "BELUM DITERIMA"
        rs_trm_tkj.Update
        rs_trm_tkj.MoveNext
    Next i
    '=============================================
    'PEMBERIAN STATUS JURUSAN SBG TANDA DITERIMA
    '=============================================
    If rs_trm_ap.RecordCount <= Right(C_trm_ap.Caption, 2) Then     'DATA KURANG DARI-
        rs_trm_ap.MoveFirst                                         '- ATO = PAGU
        For i = 1 To rs_trm_ap.RecordCount
            rs_trm_ap(7) = "AP"
            rs_trm_ap.Update
            rs_trm_ap.MoveNext
        Next i
    Else
        rs_trm_ap.MoveFirst
        For i = 1 To Right(C_trm_ap.Caption, 2)         'DATA MELEBIHI PAGU
            rs_trm_ap(7) = "AP"
            rs_trm_ap.Update
            rs_trm_ap.MoveNext
        Next i
    End If
    '=============================================AP
    If rs_trm_ak.RecordCount <= Right(C_trm_ak.Caption, 2) Then     'DATA KURANG DARI-
        rs_trm_ak.MoveFirst                                         '- ATO = PAGU
        For i = 1 To rs_trm_ak.RecordCount
            rs_trm_ak(7) = "AK"
            rs_trm_ak.Update
            rs_trm_ak.MoveNext
        Next i
    Else
        rs_trm_ak.MoveFirst
        For i = 1 To Right(C_trm_ak.Caption, 2)         'DATA MELEBIHI PAGU
            rs_trm_ak(7) = "AK"
            rs_trm_ak.Update
            rs_trm_ak.MoveNext
        Next i
    End If
    '=============================================AK
    If rs_trm_pj.RecordCount <= Right(C_trm_pj.Caption, 2) Then     'DATA KURANG DARI-
        rs_trm_pj.MoveFirst                                         '- ATO = PAGU
        For i = 1 To rs_trm_pj.RecordCount
            rs_trm_pj(7) = "PJ"
            rs_trm_pj.Update
            rs_trm_pj.MoveNext
        Next i
    Else
        rs_trm_pj.MoveFirst
        For i = 1 To Right(C_trm_pj.Caption, 2)         'DATA MELEBIHI PAGU
            rs_trm_pj(7) = "PJ"
            rs_trm_pj.Update
            rs_trm_pj.MoveNext
        Next i
    End If
    '=============================================PJ
    If rs_trm_tkj.RecordCount <= Right(C_trm_tkj.Caption, 2) Then     'DATA KURANG DARI-
        rs_trm_tkj.MoveFirst                                         '- ATO = PAGU
        For i = 1 To rs_trm_tkj.RecordCount
            rs_trm_tkj(7) = "TKJ"
            rs_trm_tkj.Update
            rs_trm_tkj.MoveNext
        Next i
    Else
        rs_trm_tkj.MoveFirst
        For i = 1 To Right(C_trm_tkj.Caption, 2)         'DATA MELEBIHI PAGU
            rs_trm_tkj(7) = "TKJ"
            rs_trm_tkj.Update
            rs_trm_tkj.MoveNext
        Next i
    End If
    '=============================================TKJ
End Sub

Private Sub KELUAR_Click()
    End
End Sub

