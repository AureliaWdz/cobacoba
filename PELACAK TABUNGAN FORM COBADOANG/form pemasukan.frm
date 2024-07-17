VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form formpemasukan 
   Caption         =   "PEMASUKAN KEUANGAN "
   ClientHeight    =   9300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11490
   LinkTopic       =   "Form1"
   ScaleHeight     =   9300
   ScaleWidth      =   11490
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   900
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
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
            Caption         =   "Pemasukan"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
            Value           =   1
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Pengeluaran"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Keluar"
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
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
      Left            =   2400
      TabIndex        =   16
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton cmdhapus 
      Caption         =   "Hapus"
      Height          =   615
      Left            =   3480
      TabIndex        =   13
      Top             =   8160
      Width           =   1695
   End
   Begin VB.TextBox txtnomor 
      Alignment       =   2  'Center
      DataField       =   "Nomor"
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
      Left            =   2400
      TabIndex        =   11
      Top             =   1800
      Width           =   975
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
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   8160
      Width           =   1575
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
      Height          =   615
      Left            =   720
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
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
      Height          =   615
      Left            =   6240
      TabIndex        =   8
      Top             =   8160
      Width           =   1455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "form pemasukan.frx":0000
      Height          =   1935
      Left            =   600
      OleObjectBlob   =   "form pemasukan.frx":0014
      TabIndex        =   7
      Top             =   5640
      Width           =   8415
   End
   Begin VB.Data Data1 
      Caption         =   "DTPEMASUKAN"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\ASUS\Documents\pemasukan\Pemasukan.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "dbmasuk"
      Top             =   5040
      Width           =   4815
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
      Height          =   405
      Left            =   2400
      TabIndex        =   6
      Top             =   4440
      Width           =   2535
   End
   Begin MSComCtl2.DTPicker dtptanggal 
      DataField       =   "Tanggal"
      DataSource      =   "Data1"
      Height          =   375
      Left            =   2400
      TabIndex        =   4
      Top             =   3840
      Width           =   2535
      _ExtentX        =   4471
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   135659520
      CurrentDate     =   45484
   End
   Begin VB.ComboBox Cmbkategori 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   3240
      Width           =   2535
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   8520
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
            Picture         =   "form pemasukan.frx":0EF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "form pemasukan.frx":120D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "form pemasukan.frx":1527
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "form pemasukan.frx":1841
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      Caption         =   "TOTAL PEMASUKAN"
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
      Left            =   7800
      TabIndex        =   15
      Top             =   1320
      Width           =   2655
   End
   Begin VB.Label Label5 
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
      Height          =   735
      Left            =   7920
      TabIndex        =   14
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label lbnomor 
      Caption         =   "Nomor :"
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
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Saldo :"
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
      TabIndex        =   5
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   720
      TabIndex        =   3
      Top             =   3960
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Kategori  : "
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
      Top             =   3240
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Rekening :"
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
      TabIndex        =   0
      Top             =   2520
      Width           =   1335
   End
End
Attribute VB_Name = "formpemasukan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdhapus_Click()
Data1.Recordset.Delete
Data1.Refresh
TotalSaldoPemasukan
End Sub
Private Sub cmdkeluar_Click()
Unload Me
End Sub
Private Sub cmdsimpan_Click()
Data1.Recordset.Update
bersih
TotalSaldoPemasukan
End Sub
Private Sub cmdtambah_Click()
Data1.Recordset.AddNew
bersih
txtnomor.SetFocus
TotalSaldoPemasukan
End Sub
Private Sub Form_Load()
cmbkategori.AddItem ("Gaji")
cmbkategori.AddItem ("Hadiah")
cmbkategori.AddItem ("Lainnya")
cmbrekening.AddItem ("Utama")
cmbrekening.AddItem ("Lainnya")
TotalSaldoPemasukan
End Sub
Sub bersih()
txtrekening = ""
cmbkategori = ""
txtnomor = ""
txtsaldo = ""
End Sub

Private Sub TotalSaldoPemasukan()
    Dim db As Database
    Dim rs As Recordset
    Dim totalSaldo As Double
    Dim saldoValue As Double
    
    ' Inisialisasi total saldo
    totalSaldo = 0
    
    ' Buka koneksi database
    Set db = OpenDatabase("C:\Users\ASUS\Documents\pemasukan\pemasukan.mdb")
    
    ' Buka recordset untuk tabel dbmasuk
    Set rs = db.OpenRecordset("SELECT saldo FROM dbmasuk")
    
    ' Loop melalui setiap record dan jumlahkan saldo
    Do While Not rs.EOF
        ' Periksa apakah nilai saldo tidak Null
        If Not IsNull(rs!saldo) Then
            ' Ambil nilai saldo dari recordset dan konversi ke Double
            saldoValue = CDbl(rs!saldo) ' Menggunakan CDbl() untuk mengonversi ke Double
            
            ' Tambahkan saldoValue ke totalSaldo
            totalSaldo = totalSaldo + saldoValue
        End If
        
        rs.MoveNext
    Loop
    
    ' Tampilkan total saldo dalam Label5
    Label5.Caption = Format(totalSaldo, "Currency")
    
    ' Tutup recordset dan koneksi database
    rs.Close
    db.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
If Button.Index = 1 Then
formberanda.Show
End If

If Button.Index = 3 Then
Formpengeluaran.Show
End If

If Button.Index = 4 Then
End
End If
End Sub
