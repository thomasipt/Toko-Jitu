VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form MAINMENU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MENU UTAMA"
   ClientHeight    =   7200
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   9450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   9450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "PEMBELIAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   7238
      TabIndex        =   11
      Top             =   1650
      Width           =   2115
   End
   Begin VB.CommandButton Command1 
      Caption         =   "PENJUALAN"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   98
      TabIndex        =   10
      Top             =   1650
      Width           =   2115
   End
   Begin VB.CommandButton cmdCLOSE 
      Caption         =   "KELUAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   98
      TabIndex        =   9
      Top             =   6585
      Width           =   9255
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6293
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3195
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   98
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6060
      Width           =   3060
   End
   Begin VB.Frame Frame3 
      Height          =   630
      Left            =   98
      TabIndex        =   2
      Top             =   5430
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   360
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0000
         TabIndex        =   4
         Top             =   165
         Width           =   8775
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1440
      Left            =   98
      TabIndex        =   0
      Top             =   0
      Width           =   9255
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0070
         TabIndex        =   5
         Top             =   945
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   450
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":00E0
         TabIndex        =   1
         Top             =   105
         Width           =   8775
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   495
         Left            =   240
         OleObjectBlob   =   "MAINMENU.frx":0150
         TabIndex        =   3
         Top             =   615
         Width           =   8775
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5520
      OleObjectBlob   =   "MAINMENU.frx":01C0
      Top             =   9900
   End
   Begin Crystal.CrystalReport Crpt 
      Left            =   525
      Top             =   9030
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   300
      Left            =   2955
      OleObjectBlob   =   "MAINMENU.frx":03F4
      TabIndex        =   12
      Top             =   1650
      Width           =   3540
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   255
      Left            =   2955
      OleObjectBlob   =   "MAINMENU.frx":0461
      TabIndex        =   13
      Top             =   2265
      Width           =   3540
   End
   Begin VB.PictureBox Picture1 
      Height          =   2775
      Left            =   98
      Picture         =   "MAINMENU.frx":04EC
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   613
      TabIndex        =   14
      Top             =   2700
      Width           =   9255
   End
   Begin VB.Menu P 
      Caption         =   "PENJUALAN"
      Index           =   1
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PENJUALAN"
         Index           =   11
      End
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PULSA"
         Index           =   12
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "-"
         Index           =   13
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "TRANSAKSI PIUTANG"
         Index           =   14
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "DAFTAR PIUTANG"
         Index           =   15
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "-"
         Index           =   16
         Visible         =   0   'False
      End
      Begin VB.Menu PJ 
         Caption         =   "CETAK ULANG NOTA"
         Index           =   17
         Visible         =   0   'False
      End
   End
   Begin VB.Menu B 
      Caption         =   "PEMBELIAN"
      Index           =   2
      Begin VB.Menu PB 
         Caption         =   "TRANSAKSI PEMBELIAN"
         Index           =   21
      End
      Begin VB.Menu PB 
         Caption         =   "DEPOSIT PULSA"
         Index           =   22
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "-"
         Index           =   23
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "HUTANG"
         Index           =   24
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "DAFTAR HUTANG"
         Index           =   25
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "-"
         Index           =   26
         Visible         =   0   'False
      End
      Begin VB.Menu PB 
         Caption         =   "RETURN"
         Index           =   27
         Visible         =   0   'False
      End
   End
   Begin VB.Menu D 
      Caption         =   "DATA"
      Index           =   31
      Begin VB.Menu DS 
         Caption         =   "KODE KATEGORI BARANG"
         Index           =   31
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "KODE BARANG"
         Index           =   32
      End
      Begin VB.Menu DS 
         Caption         =   "KODE PELANGGAN"
         Index           =   33
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "KODE DISTRIBUTOR"
         Index           =   34
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "-"
         Index           =   35
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "JASA"
         Index           =   36
         Visible         =   0   'False
      End
      Begin VB.Menu DS 
         Caption         =   "PULSA ELEKTRONIK"
         Index           =   37
         Visible         =   0   'False
      End
   End
   Begin VB.Menu T 
      Caption         =   "TOOLS"
      Index           =   4
      Begin VB.Menu TS 
         Caption         =   "SETING TOKO"
         Index           =   41
      End
      Begin VB.Menu TS 
         Caption         =   "GANTI PASSWORD"
         Index           =   42
      End
      Begin VB.Menu TS 
         Caption         =   "USER BARU"
         Index           =   43
         Visible         =   0   'False
      End
   End
   Begin VB.Menu L 
      Caption         =   "LAPORAN"
      Index           =   5
      Begin VB.Menu LS 
         Caption         =   "MUTASI BARANG"
         Index           =   501
      End
      Begin VB.Menu LS 
         Caption         =   "STOCK LIMIT"
         Index           =   502
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   503
      End
      Begin VB.Menu LS 
         Caption         =   "LAP PEMBELIAN"
         Index           =   504
      End
      Begin VB.Menu LS 
         Caption         =   "LAP JUMLAH PEMBELIAN"
         Index           =   505
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   506
      End
      Begin VB.Menu LS 
         Caption         =   "LAP PENJUALAN"
         Index           =   507
      End
      Begin VB.Menu LS 
         Caption         =   "LAP JUMLAH PENJUALAN"
         Index           =   508
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "-"
         Index           =   509
         Visible         =   0   'False
      End
      Begin VB.Menu LS 
         Caption         =   "LABA RUGI"
         Index           =   510
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "MASUK"
         Index           =   60
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "KELUAR"
         Index           =   61
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "-"
         Index           =   62
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "DAFTAR SERVICE"
         Index           =   63
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "-"
         Index           =   64
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "NOTA SERVICE MASUK"
         Index           =   65
         Visible         =   0   'False
      End
      Begin VB.Menu SS 
         Caption         =   "NOTA SERVICE KELUAR"
         Index           =   66
         Visible         =   0   'False
      End
   End
   Begin VB.Menu A 
      Caption         =   "ABOUT"
      Index           =   7
      Begin VB.Menu AA 
         Caption         =   "LISENSI"
         Index           =   70
      End
   End
End
Attribute VB_Name = "MAINMENU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String

Private Sub AA_Click(Index As Integer)
Select Case Index
    Case 70
        L001.Show 1
End Select
End Sub

Private Sub cmdCLOSE_Click()
Unload Me
LOGIN.Show
End Sub

Private Sub Command1_Click()
JL001.Show 1
End Sub

Private Sub Command2_Click()
BL001.Show 1
End Sub

Private Sub DS_Click(Index As Integer)
Select Case Index
    Case 31
        B001.Show 1
    Case 32
        B003.Show 1
    Case 33
        P001.Show 1
    Case 34
        D001.Show 1
    Case 36
        JS01.Show 1
    Case 37
        VC01.Show 1
End Select
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd

Text1 = "USER : " + Operator
Text2 = Date
Text3 = "Copyrighted 2008 - EDP_IPT"

SkinLabel1 = NTOKO
SkinLabel4 = NAlamat
SkinLabel5 = NMOtto
SkinLabel6 = NTelepon
'Me.Left = 0
'Me.Top = 0
End Sub

Private Sub LS_Click(Index As Integer)
Select Case Index
    Case 501
        Call LapBR
    Case 504
        Indikator = 0
        TglFuck = ""
        TGLFAK.Show 1
        If Indikator = 1 Then
            Call LapTransBeli
        Else
            Exit Sub
        End If
    Case 505
    Case 506
    Case 507
        Indikator = 0
        TglFuck = ""
        TglFuck2 = ""
        TGLFAK.Show 1
        If Indikator = 1 Then
            Call LapTransJual
        Else
            Exit Sub
        End If
    Case 508
    Case 509
    Case 510
        LR.Show 1
End Select
End Sub

Private Sub LapBR()
crpt.ReportFileName = "C:\WINDOWS\Reporttoko\LapBR.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub LapTransBeli()
crpt.ReportFileName = "C:\WINDOWS\Reporttoko\TransBeli.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub LapTransJual()
crpt.ReportFileName = "C:\WINDOWS\Reporttoko\TransJual.rpt"
crpt.WindowState = crptMaximized
crpt.WindowMaxButton = False
crpt.WindowMinButton = False
crpt.Action = 1
End Sub

Private Sub PB_Click(Index As Integer)
Select Case Index
    Case 21
        BL001.Show 1
    Case 22
        VC02.Show 1
    Case 24
    Case 25

End Select
End Sub

Private Sub PJ_Click(Index As Integer)
Select Case Index
    Case 11
        JL001.Show 1
    Case 12
        VC03.Show 1
    Case 14
    Case 15
End Select
End Sub

Private Sub SS_Click(Index As Integer)
Select Case Index
    Case 60
        JS02.Show 1
    Case 61
        JS03.Show 1
    Case 63
        JS001.Show 1
End Select
End Sub

Private Sub TS_Click(Index As Integer)
Select Case Index
    Case 41
        NAMA.Show 1
    Case 42
        GPASS.Show 1
    Case 43
        User.Show 1
End Select
End Sub
