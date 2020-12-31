VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "ENTRY DATA"
   ClientHeight    =   8190
   ClientLeft      =   60
   ClientTop       =   -1710
   ClientWidth     =   9345
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9345
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin VB.TextBox T_juml_data 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   14280
      TabIndex        =   60
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmd_cabut 
      Caption         =   "PENCABUTAN DATA"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10200
      TabIndex        =   58
      Top             =   1560
      Width           =   1935
   End
   Begin VB.TextBox txt_noDaftar 
      Height          =   375
      Left            =   8880
      TabIndex        =   57
      Top             =   1560
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7695
      Left            =   7200
      TabIndex        =   56
      Top             =   2280
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   13573
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
   Begin VB.TextBox T_PKN 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   49
      Text            =   "0"
      Top             =   6720
      Width           =   2295
   End
   Begin VB.TextBox T_IPS 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   48
      Text            =   "0"
      Top             =   6360
      Width           =   2295
   End
   Begin VB.TextBox T_IPA 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   47
      Text            =   "0"
      Top             =   6000
      Width           =   2295
   End
   Begin VB.ComboBox C_pil_1 
      Height          =   315
      ItemData        =   "Form entri data.frx":0000
      Left            =   1560
      List            =   "Form entri data.frx":0010
      TabIndex        =   13
      Top             =   8400
      Width           =   5415
   End
   Begin VB.ComboBox C_pil_2 
      Height          =   315
      ItemData        =   "Form entri data.frx":0025
      Left            =   1560
      List            =   "Form entri data.frx":0035
      TabIndex        =   12
      Top             =   8760
      Width           =   5415
   End
   Begin VB.TextBox T_nodaftar 
      Height          =   405
      Left            =   2280
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
   End
   Begin VB.TextBox T_nama_siswa 
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      Top             =   1560
      Width           =   4695
   End
   Begin VB.TextBox T_lahir 
      Height          =   375
      Left            =   2280
      TabIndex        =   9
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox T_nama_ortu 
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox T_alamat 
      Height          =   375
      Left            =   2280
      TabIndex        =   7
      Top             =   2640
      Width           =   4695
   End
   Begin VB.TextBox T_krj_ortu 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   4695
   End
   Begin VB.TextBox T_smp_asal 
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   3360
      Width           =   4695
   End
   Begin VB.TextBox T_bin 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Text            =   "0"
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox T_big 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   3
      Text            =   "0"
      Top             =   5280
      Width           =   2295
   End
   Begin VB.TextBox T_mat 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1800
      TabIndex        =   2
      Text            =   "0"
      Top             =   5640
      Width           =   2295
   End
   Begin VB.CommandButton C_entry 
      Caption         =   "ENTRY"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   9360
      Width           =   3375
   End
   Begin VB.CommandButton C_clear 
      Caption         =   "CLEA&R"
      Height          =   495
      Left            =   3600
      TabIndex        =   0
      Top             =   9360
      Width           =   3375
   End
   Begin VB.Label Label27 
      Caption         =   "Data Sementara"
      Height          =   255
      Left            =   12840
      TabIndex        =   61
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label26 
      Caption         =   "Nomor Pendaftaran"
      Height          =   255
      Left            =   7200
      TabIndex        =   59
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label L_bobot_ips 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   55
      Top             =   6480
      Width           =   855
   End
   Begin VB.Label L_bobot_ipa 
      Alignment       =   2  'Center
      Caption         =   "2"
      Height          =   255
      Left            =   4320
      TabIndex        =   54
      Top             =   6120
      Width           =   855
   End
   Begin VB.Label L_bobot_pkn 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   53
      Top             =   6840
      Width           =   855
   End
   Begin VB.Label L_nun_x_bbt_pkn 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   52
      Top             =   6840
      Width           =   1095
   End
   Begin VB.Label L_nun_x_bbt_ips 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   51
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Label L_nun_x_bbt_ipa 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   50
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Line Line9 
      X1              =   120
      X2              =   6960
      Y1              =   7680
      Y2              =   7680
   End
   Begin VB.Line Line7 
      X1              =   120
      X2              =   6960
      Y1              =   7200
      Y2              =   7200
   End
   Begin VB.Line Line6 
      X1              =   5280
      X2              =   5280
      Y1              =   4440
      Y2              =   7680
   End
   Begin VB.Line Line4 
      X1              =   4200
      X2              =   4200
      Y1              =   4440
      Y2              =   7680
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   1680
      Y1              =   4440
      Y2              =   7680
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   6960
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Label Label25 
      Caption         =   "PPKN"
      Height          =   255
      Left            =   240
      TabIndex        =   46
      Top             =   6840
      Width           =   1335
   End
   Begin VB.Label Label24 
      Caption         =   "IPS"
      Height          =   255
      Left            =   240
      TabIndex        =   45
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label23 
      Caption         =   "IPA"
      Height          =   255
      Left            =   240
      TabIndex        =   44
      Top             =   6120
      Width           =   1335
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   6960
      Y1              =   4800
      Y2              =   4800
   End
   Begin VB.Label Label20 
      Caption         =   "PENCARIAN/PENCABUTAN DATA"
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
      Left            =   7200
      TabIndex        =   43
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label Label19 
      Caption         =   "NILAI DAN PEMBOBOTAN                                       (tanda koma tergantung settingan bhs)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   42
      Top             =   3840
      Width           =   6855
   End
   Begin VB.Label Label2 
      Caption         =   "Nama siswa"
      Height          =   255
      Left            =   120
      TabIndex        =   41
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Tempat dan tanggal lahir"
      Height          =   255
      Left            =   120
      TabIndex        =   40
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Nama Orang Tua"
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   2400
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Alamat"
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Label7 
      Caption         =   "Pekerjaan Orang Tua"
      Height          =   255
      Left            =   120
      TabIndex        =   37
      Top             =   3120
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "SMP asal"
      Height          =   255
      Left            =   120
      TabIndex        =   36
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Caption         =   "NILAI YANG DISYARATKAN"
      Height          =   375
      Left            =   1800
      TabIndex        =   35
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label10 
      Caption         =   "Bahasa Indonesia"
      Height          =   255
      Left            =   240
      TabIndex        =   34
      Top             =   5040
      Width           =   1335
   End
   Begin VB.Label Label11 
      Caption         =   "Bahasa Inggris"
      Height          =   255
      Left            =   240
      TabIndex        =   33
      Top             =   5400
      Width           =   1335
   End
   Begin VB.Label Label12 
      Caption         =   "Matematika"
      Height          =   255
      Left            =   240
      TabIndex        =   32
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Label Label13 
      Caption         =   "PILIHAN JURUSAN"
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
      TabIndex        =   31
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Label Label14 
      Caption         =   "DATA PRIBADI"
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
      TabIndex        =   30
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label15 
      Caption         =   "Pilihan I"
      Height          =   255
      Left            =   120
      TabIndex        =   29
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Label Label16 
      Caption         =   "Pilihan II"
      Height          =   255
      Left            =   120
      TabIndex        =   28
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "No Pendaftaran"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "JUMLAH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label L_jml_nun 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   7320
      Width           =   1335
   End
   Begin VB.Label L_bobot_mat 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Left            =   4320
      TabIndex        =   24
      Top             =   5760
      Width           =   855
   End
   Begin VB.Label L_bobot_bin 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Left            =   4320
      TabIndex        =   23
      Top             =   5040
      Width           =   855
   End
   Begin VB.Label L_bobot_big 
      Alignment       =   2  'Center
      Caption         =   "3"
      Height          =   255
      Left            =   4320
      TabIndex        =   22
      Top             =   5400
      Width           =   855
   End
   Begin VB.Label Label21 
      Alignment       =   2  'Center
      Caption         =   "BOBOT"
      Height          =   255
      Left            =   4320
      TabIndex        =   21
      Top             =   4560
      Width           =   855
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      Caption         =   "NILAI  X  BOBOT"
      Height          =   255
      Left            =   5400
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label L_nun_x_bbt_bin 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   19
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label L_nun_x_bbt_big 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   18
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label L_nun_x_bbt_mat 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   17
      Top             =   5760
      Width           =   1095
   End
   Begin VB.Label L_TOTAL_BOBOT 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   5520
      TabIndex        =   16
      Top             =   7320
      Width           =   1095
   End
   Begin VB.Line Line5 
      X1              =   6960
      X2              =   6960
      Y1              =   4440
      Y2              =   7680
   End
   Begin VB.Line Line8 
      X1              =   120
      X2              =   120
      Y1              =   4440
      Y2              =   7680
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      Caption         =   "PENDAFTARAN PESERTA DIDIK BARU (PPDB) SMKN 1 NGAWI"
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
      TabIndex        =   15
      Top             =   120
      Width           =   15015
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      Caption         =   "PELAJARAN"
      Height          =   255
      Left            =   240
      TabIndex        =   14
      Top             =   4560
      Width           =   1335
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset




Private Sub C_clear_Click()
    Form1.T_nodaftar.Text = ""
    Form1.T_nama_siswa.Text = ""
    Form1.T_lahir.Text = ""
    Form1.T_nama_ortu.Text = ""
    Form1.T_alamat.Text = ""
    Form1.T_krj_ortu.Text = ""
    Form1.T_smp_asal.Text = ""
    Form1.C_pil_1.Text = ""
    Form1.C_pil_2.Text = ""
    txt_noDaftar.Text = ""
    cmd_cabut.Enabled = False
    
    T_bin.Text = 0
    T_big.Text = 0
    T_mat.Text = 0
    T_IPA.Text = 0
    T_IPS.Text = 0
    T_PKN.Text = 0
    L_nun_x_bbt_bin.Caption = 0
    L_nun_x_bbt_big.Caption = 0
    L_nun_x_bbt_mat.Caption = 0
    L_nun_x_bbt_ipa.Caption = 0
    L_nun_x_bbt_ips.Caption = 0
    L_nun_x_bbt_pkn.Caption = 0
    L_jml_nun.Caption = 0
    L_TOTAL_BOBOT.Caption = 0
    
    Form1.T_nodaftar.SetFocus
End Sub

Private Sub C_entry_Click()
    rs.AddNew
    rs(0) = Form1.T_nodaftar.Text
    rs(1) = Form1.T_nama_siswa.Text
    rs(2) = Form1.T_lahir.Text
    rs(3) = Form1.T_nama_ortu.Text
    rs(4) = Form1.T_alamat.Text
    rs(5) = Form1.T_krj_ortu.Text
    rs(6) = Form1.T_smp_asal.Text
    rs(7) = CSng(Form1.T_bin.Text)
    rs(8) = CSng(Form1.T_big.Text)
    rs(9) = CSng(Form1.T_mat.Text)
    rs(10) = CSng(Form1.T_IPA.Text)
    rs(11) = CSng(Form1.T_IPS.Text)
    rs(12) = CSng(Form1.T_PKN.Text)
    rs(13) = Form1.C_pil_1.Text
    rs(14) = Form1.C_pil_2.Text
    rs(15) = "TERDAFTAR"
    rs(16) = Date
    rs(17) = Form1.L_TOTAL_BOBOT.Caption
    rs.Update
    
    C_clear_Click
    
    rs.Requery
    rs.Sort = "no_daftar asc"
    rs.MoveLast
    Set DataGrid1.DataSource = rs
End Sub

Private Sub CABUT_DATA_Click()
    Form1.Hide
    Form2.Show
End Sub

Private Sub cmd_cabut_Click()
    cmd_cabut.Enabled = False
    
    rs(15) = "DICABUT"
    rs.Update
    txt_noDaftar.Text = ""
    txt_noDaftar.SetFocus
End Sub

Public Sub Form_Load()
    '=============================================
    'MEMBUKA KONEKSI DATABASE
    '=============================================
    If Form1.con.State = 1 Then Form1.con.Close
    con.CursorLocation = adUseClient
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\dsa\data calon peserta didik.mdb;"
    '=============================================
    'MENYIMPAN SELURUH RECORD DI VARIABEL rs
    '=============================================
    rs.Open "select * from table1", con, adOpenKeyset, adLockOptimistic
    '=============================================
    'MENGURUTKAN RECORD PADA FIELD NO_DAFTAR
    '=============================================
    rs.Sort = "no_daftar asc"
    rs.MoveLast
    '=============================================
    'MENAMPILKAN RECORD DI DATAGRID1
    '=============================================
    Set DataGrid1.DataSource = rs
    '=============================================
    'MENGHITUNG JUMLAH RECORD
    '=============================================
    T_juml_data.Text = rs.RecordCount
End Sub

Private Sub KELUAR_Click()
    
    End
End Sub

Private Sub PROSES_DATA_Click()
'    rs_trm_ap.Close
    Form1.Hide
    Form3.Form_Load
    Form3.Show
'    Form3.cmd_proses_lanjutan.Caption = "PROSES KE 3"
End Sub


Private Sub T_alamat_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_krj_ortu.SetFocus
    End If
End Sub

Private Sub T_big_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        Form1.T_mat.Text = ""
        Form1.T_mat.SetFocus
    End If
End Sub

Private Sub T_bin_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        Form1.T_big.Text = ""
        Form1.T_big.SetFocus
    End If
End Sub

Private Sub T_IPA_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        Form1.T_IPS.Text = ""
        Form1.T_IPS.SetFocus
    End If
End Sub

Private Sub T_IPS_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        Form1.T_PKN.Text = ""
        Form1.T_PKN.SetFocus
    End If
End Sub

Private Sub T_krj_ortu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_smp_asal.SetFocus
    End If
End Sub

Private Sub T_lahir_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_nama_ortu.SetFocus
    End If
End Sub

Private Sub T_mat_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        T_IPA.Text = ""
        T_IPA.SetFocus
    End If
End Sub

Private Sub T_nama_ortu_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_alamat.SetFocus
    End If
End Sub

Private Sub T_nama_siswa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_lahir.SetFocus
    End If
End Sub

Private Sub T_nodaftar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_nama_siswa.SetFocus
'    elseif KeyAscii=
    End If
End Sub

Private Sub T_PKN_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        perkalian_bobot_nilai
        
        C_pil_1.SetFocus
    End If
End Sub

Private Sub T_smp_asal_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        T_bin.Text = ""
        T_bin.SetFocus
    End If
End Sub

Private Sub C_pil_1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        C_pil_2.SetFocus
    End If
End Sub

Private Sub C_pil_2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        C_entry.SetFocus
    End If
End Sub

Function perkalian_bobot_nilai()
    '=========================================
    'INISIALISASI
    '=========================================
    Form1.L_nun_x_bbt_bin.Caption = 0
    Form1.L_nun_x_bbt_big.Caption = 0
    Form1.L_nun_x_bbt_mat.Caption = 0
    Form1.L_nun_x_bbt_ipa.Caption = 0
    Form1.L_nun_x_bbt_ips.Caption = 0
    Form1.L_nun_x_bbt_pkn.Caption = 0
    
    Form1.L_jml_nun.Caption = ""
    Form1.L_TOTAL_BOBOT.Caption = 0
    '=========================================
    'PROSES PERKALIAN
    '=========================================
    Form1.L_nun_x_bbt_bin.Caption = CSng(Form1.T_bin.Text) * CSng(Form1.L_bobot_bin.Caption)
    Form1.L_nun_x_bbt_big.Caption = CSng(Form1.T_big.Text) * CSng(Form1.L_bobot_big.Caption)
    Form1.L_nun_x_bbt_mat.Caption = CSng(Form1.T_mat.Text) * CSng(Form1.L_bobot_mat.Caption)
    Form1.L_nun_x_bbt_ipa.Caption = CSng(Form1.T_IPA.Text) * CSng(Form1.L_bobot_ipa.Caption)
    Form1.L_nun_x_bbt_ips.Caption = CSng(Form1.T_IPS.Text) * CSng(Form1.L_bobot_ips.Caption)
    Form1.L_nun_x_bbt_pkn.Caption = CSng(Form1.T_PKN.Text) * CSng(Form1.L_bobot_pkn.Caption)
    
    Form1.L_jml_nun.Caption = CSng(Form1.T_bin.Text) + CSng(Form1.T_big.Text) + CSng(Form1.T_mat.Text) + CSng(Form1.T_IPA.Text) + CSng(Form1.T_IPS.Text) + CSng(Form1.T_PKN.Text)
    Form1.L_TOTAL_BOBOT.Caption = CSng(Form1.L_nun_x_bbt_bin.Caption) + CSng(Form1.L_nun_x_bbt_big.Caption) + CSng(Form1.L_nun_x_bbt_mat.Caption) + CSng(Form1.L_nun_x_bbt_ipa.Caption) + CSng(Form1.L_nun_x_bbt_ips.Caption) + CSng(Form1.L_nun_x_bbt_pkn.Caption)
End Function

Private Sub txt_noDaftar_KeyPress(KeyAscii As Integer)
    On Error Resume Next
    If KeyAscii = 13 Then
        rs.Sort = "no_daftar asc"
        rs.MoveFirst
        rs.Find "no_daftar='" & txt_noDaftar.Text & "'"
        If rs.EOF Then
            MsgBox "data tidak ditemukan"
            cmd_cabut.Enabled = False
            txt_noDaftar.SetFocus
            txt_noDaftar.Text = ""
        ElseIf rs(15) = "DICABUT" Then
            MsgBox "data data telah dicabut"
            cmd_cabut.Enabled = False
            txt_noDaftar.SetFocus
            txt_noDaftar.Text = ""
        Else
            cmd_cabut.Enabled = True
            cmd_cabut.SetFocus
        End If
        
        
'        On Error Resume Next
'        With Adodc1.Recordset
'            .MoveFirst
'            .Find "no_daftar='" & txt_noDaftar.Text & "'"
'        End With
    End If
End Sub
