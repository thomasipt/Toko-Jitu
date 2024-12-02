VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form VC02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSAKSI DEPOSIT PULSA"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
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
      Left            =   2070
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   2040
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
      Left            =   5055
      TabIndex        =   6
      Top             =   1590
      Width           =   1890
   End
   Begin VB.CommandButton cmdDEL 
      Caption         =   "HAPUS"
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
      Left            =   2662
      TabIndex        =   5
      Top             =   1590
      Width           =   1890
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2055
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   990
      Width           =   2490
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   2055
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   585
      Width           =   2490
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6615
      OleObjectBlob   =   "VC02.frx":0000
      Top             =   720
   End
   Begin MSFlexGridLib.MSFlexGrid grid 
      Height          =   2220
      Left            =   97
      TabIndex        =   7
      Top             =   2430
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   3916
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   65280
      BackColorBkg    =   16777152
      GridColor       =   0
      Enabled         =   -1  'True
      TextStyle       =   3
      TextStyleFixed  =   3
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0234
      TabIndex        =   8
      Top             =   180
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   4185
      OleObjectBlob   =   "VC02.frx":029A
      TabIndex        =   9
      Top             =   195
      Width           =   2955
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0300
      TabIndex        =   10
      Top             =   1050
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "VC02.frx":0372
      TabIndex        =   12
      Top             =   645
      Width           =   1695
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "SIMPAN"
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
      Left            =   270
      TabIndex        =   3
      Top             =   1590
      Width           =   1890
   End
   Begin VB.CommandButton cmdEDIT 
      Caption         =   "EDIT"
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
      Left            =   270
      TabIndex        =   4
      Top             =   1590
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   -270
      ScaleHeight     =   795
      ScaleWidth      =   8145
      TabIndex        =   11
      Top             =   1440
      Width           =   8205
   End
End
Attribute VB_Name = "VC02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi As String

Private RDOE As rdoEnvironment
Private RDCO As rdoConnection
Private RSLNO As rdoResultset

Private RSL, RSLUser, RCari, RCari2, RCari3, RCari4, RCari5, RSave, RSave2, RSave3, RSave4, RSave5, REdit As rdoResultset
Private SQL, SQLUser, SCari, SCari2, SCari3, SCari4, SCari5, SSave, SSave2, SSave3, SSave4, SSave5, SEdit As String

Private RJual1, RJual2, RJual3, RJual4, RJual5, RJual6, RJual7, RJual8, RJual9, RJual10 As rdoResultset
Private SJual1, SJual2, SJual3, SJual4, SJual5, SJual6, SJual7, SJual8, SJual9, SJual10 As String

Private RBahan1, RBahan2, RBahan3, RBahan4, RBahan5, RBahan6, RBahan7, RBahan8, RBahan9, RBahan10 As rdoResultset
Private SBahan1, SBahan2, SBahan3, SBahan4, SBahan5, SBahan6, SBahan7, SBahan8, SBahan9, SBahan10 As String

Private RDEl As rdoResultset
Private SDel As String

Private RLR, RLR2 As rdoResultset
Private SLR, SLR2 As String

Private RJS As rdoResultset
Private SJS As String

Private SqlNo As String

Private Sub cmdCLOSE_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Combo1 = "" Or Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

Call Simpan
Call Simpan2

Unload Me
VC02.Show 1
End Sub

Private Sub Simpan()
SSave = "Select * From VC02"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("Tanggal") = Date
        RSave("Kode") = Trim(Combo1)
        RSave("Nama") = Trim(SkinLabel3)
        RSave("Stok") = CCur(Text1)
        RSave("Beli") = CCur(Text2)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub Simpan2()
Dim Pusing As String

SSave2 = "Select * From VC01 where Kode = '" + Trim(Combo1) + "'"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    Pusing = CCur(RSave2("Satuan"))
RSave2.Edit
    RSave2("Beli") = CCur(Text1)
    RSave2("Stok") = CCur(Text2)
    RSave2("Stokbel") = CCur(Text1)
    RSave2("Jumlah") = CCur(Text1) / CCur(Pusing)
RSave2.Update
RSave2.Close
Set RSave2 = Nothing
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{TAB}"
End If
End Sub

Private Sub Combo1_LostFocus()
If Combo1 = "" Then Exit Sub
SCari2 = "Select * From VC01 where Kode = '" + Combo1 + "'"
Set RCari2 = RDCO.OpenResultset(SCari2, rdOpenDynamic, rdConcurRowVer)
If RCari2.RowCount <> 0 Then
    SkinLabel3 = RCari2("Nama")
Else
    MsgBox "KODE BELUM TERDAFTAR", vbCritical, "KONFIRMASI"
    Combo1.SetFocus
End If
RCari2.Close
Set RCari2 = Nothing
End Sub

Private Sub Form_Load()
Lokasi = App.Path
Skin1.LoadSkin Lokasi + "\" + Skin + ".skn"
Skin1.ApplySkin hWnd
Set RDOE = rdoEnvironments(0)
Set RDCO = RDOE.OpenConnection("DSN=SELULER", rdDriverNoPrompt, False, CN)

ClearTextBoxes Me

On Error GoTo ErrorHandler
SSPL = "Select Kode From VC01 order by KODE"
Set RSPL = RDCO.OpenResultset(SSPL, rdOpenDynamic, rdOpenKeyset)
RSPL.MoveFirst
Do While Not RSPL.EOF
    Combo1.AddItem RSPL("Kode")
RSPL.MoveNext
Loop
RSPL.Close
Set RSPL = Nothing
Combo1.ListIndex = 0

SkinLabel3 = ""

Call SiapkanGrid
Call IsiGrid

cmdOK.Visible = True
cmdEDIT.Visible = False
cmdDEL.Visible = False

ErrorHandler:
Select Case Err.Number
    Case 380
    Combo1 = ""
    SkinLabel3 = ""
    MsgBox "DATA KODE MASIH KOSONG", vbCritical, "KONFIRMASI"
End Select

End Sub

Private Sub SiapkanGrid()
With grid
    .Row = 0
    .Cols = 4
    .Col = 0: .ColWidth(0) = 1000: .Text = "KODE": .CellAlignment = 4
    .Col = 1: .ColWidth(1) = 2000: .Text = "NAMA": .CellAlignment = 4
    .Col = 2: .ColWidth(2) = 1500: .Text = "JUMLAH": .CellAlignment = 4
    .Col = 3: .ColWidth(3) = 1500: .Text = "H. BELI": .CellAlignment = 4
End With
End Sub

Private Sub IsiGrid()
SCari = "Select * From VC01 order by KODE"
Set RCari = RDCO.OpenResultset(SCari, rdOpenKeyset, rdConcurReadOnly)
If RCari.RowCount <> 0 Then
   RCari.MoveFirst
   B = 1
   Do Until RCari.EOF
      grid.Rows = B + 1
      grid.Row = B
         With grid
              .Col = 0: .Text = RCari("Kode"): .CellAlignment = 4
              .Col = 1: .Text = RCari("Nama")
              .Col = 2: .Text = Format(RCari("Stok"), "##,###.00")
              .Col = 3: .Text = Format(RCari("Beli"), "##,###.00")
         End With
      B = B + 1
      RCari.MoveNext
   Loop
End If
RCari.Close
Set RCari = Nothing
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text1_LostFocus()
Text1 = Format(Text1, "##,###.00")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not (KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = vbKeyBack Or KeyAscii = vbKeyReturn) Then
KeyAscii = 0
End If
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Text2_LostFocus()
Text2 = Format(Text2, "##,###.00")
End Sub
