VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Formpengeluaran 
   Caption         =   "PENGELUARAN KEUANGAN"
   ClientHeight    =   9720
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13950
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   13950
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   13950
      _ExtentX        =   24606
      _ExtentY        =   1588
      ButtonWidth     =   1931
      ButtonHeight    =   1429
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Beranda"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pemasukan "
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pengeluaran"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
            Value           =   1
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Keluar"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "fompengeluaran.frx":0000
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "fompengeluaran.frx":0014
      TabIndex        =   16
      Top             =   6600
      Width           =   13695
   End
   Begin VB.Data Data1 
      Caption         =   "Pengeluaran"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\ASUS\Documents\pemasukan\Pengeluaran.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Pengeluaran"
      Top             =   8400
      Width           =   3015
   End
   Begin VB.CommandButton cmdkeluar 
      Caption         =   "Keluar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   11400
      TabIndex        =   15
      Top             =   8640
      Width           =   1815
   End
   Begin VB.CommandButton cmdsimpan 
      Caption         =   "Simpan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8520
      TabIndex        =   14
      Top             =   8640
      Width           =   1935
   End
   Begin VB.ComboBox cmbkategori 
      DataField       =   "Kategori"
      DataSource      =   "Data1"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   4800
      Width           =   2295
   End
   Begin MSComCtl2.DTPicker dtptanggal 
      DataField       =   "Tanggal"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2280
      TabIndex        =   12
      Top             =   3960
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   135790592
      CurrentDate     =   45485
   End
   Begin VB.ComboBox cmbrekening 
      DataField       =   "Rekening"
      DataSource      =   "Data1"
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
      Left            =   2280
      TabIndex        =   11
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox txtsaldo 
      DataField       =   "Saldo"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   5760
      Width           =   3855
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5160
      TabIndex        =   9
      Top             =   8640
      Width           =   2055
   End
   Begin VB.TextBox txtkomentar 
      DataField       =   "Komentar"
      DataSource      =   "Data1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8040
      TabIndex        =   8
      Top             =   4680
      Width           =   4455
   End
   Begin VB.TextBox txtnomor 
      DataField       =   "Nomor"
      DataSource      =   "Data1"
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdtambah 
      Caption         =   "Tambah"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   480
      TabIndex        =   0
      Top             =   1200
      Width           =   1695
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   12840
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   4
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fompengeluaran.frx":09E7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fompengeluaran.frx":0D01
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fompengeluaran.frx":101B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "fompengeluaran.frx":1335
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8520
      TabIndex        =   18
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Label Label1 
      Caption         =   "TOTAL PENGELUARAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   2160
      Width           =   2535
   End
   Begin VB.Label lbsaldo 
      Caption         =   "Saldo : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label lbkomentar 
      Alignment       =   2  'Center
      Caption         =   "Komentar :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   5
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lbtanggal 
      Caption         =   "Tanggal : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Label lbkategori 
      Caption         =   "Kategori : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   4800
      Width           =   1215
   End
   Begin VB.Label lbrekening 
      Caption         =   " Rekening :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Label lbnomor 
      Caption         =   "Nomor : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2400
      Width           =   855
   End
End
Attribute VB_Name = "Formpengeluaran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdhapus_Click()
Data1.Recordset.Delete
Data1.Refresh
totalSaldoPengeluaran
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
Data1.Recordset.Update
bersih
totalSaldoPengeluaran
End Sub
Private Sub cmdtambah_Click()
Data1.Recordset.AddNew
bersih
txtnomor.SetFocus
End Sub
Private Sub Form_Load()
cmbkategori.AddItem ("Kesehatan")
cmbkategori.AddItem ("Hiburan")
cmbkategori.AddItem ("Pendidikan")
cmbkategori.AddItem ("Keperluan rumah")
cmbkategori.AddItem ("Hadiah")
cmbkategori.AddItem ("Bahan Makanan")
cmbkategori.AddItem ("Transportasi")
cmbkategori.AddItem ("Skincare dan Makeup")
cmbkategori.AddItem ("Olahraga")
cmbkategori.AddItem ("Jajan")
cmbkategori.AddItem ("Lainnya")
cmbrekening.AddItem ("Utama")
cmbrekening.AddItem ("Lainnya")
totalSaldoPengeluaran
End Sub

Sub bersih()
cmbrekening = ""
cmbkategori = ""
txtnomor = ""
txtsaldo = ""
txtkomentar = ""
End Sub
Private Sub totalSaldoPengeluaran()
    Dim db As Database
    Dim rs As Recordset
    Dim totalSaldo As Double
    
    ' Inisialisasi total saldo
    totalSaldo = 0
    
    On Error GoTo ErrorHandler
    
    ' Buka koneksi database menggunakan DAO
    Set db = OpenDatabase("C:\Users\ASUS\Documents\pemasukan\pengeluaran.mdb")
    
    ' Buka recordset untuk tabel pengeluaran
    Set rs = db.OpenRecordset("SELECT saldo FROM pengeluaran")
    
    ' Loop melalui setiap record dan jumlahkan saldo
    Do While Not rs.EOF
        ' Konversi nilai saldo dari Text ke Double
        If IsNumeric(rs!saldo) Then
            totalSaldo = totalSaldo + CDbl(rs!saldo)
        End If
        rs.MoveNext
    Loop
    
    ' Tampilkan total saldo dalam Label
    Label2.Caption = Format(totalSaldo, "Currency")
    
    ' Tutup recordset dan koneksi database
    rs.Close
    db.Close
    
    ' Membersihkan objek
    Set rs = Nothing
    Set db = Nothing
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation
    Resume Next
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Index = 1 Then
formberanda.Show
End If

If Button.Index = 2 Then
formpemasukan.Show
End If

If Button.Index = 4 Then
End
End If
End Sub
