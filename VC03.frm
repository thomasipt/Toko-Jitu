VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form VC03 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PENJUALAN PULSA ELETRONIK"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
      Height          =   240
      Left            =   180
      OleObjectBlob   =   "VC03.frx":0000
      TabIndex        =   14
      Top             =   3600
      Width           =   2955
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   45
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   3465
      Width           =   7080
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
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
      Left            =   2010
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1410
      Width           =   1095
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
      Left            =   2010
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1950
      Width           =   5010
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
      Left            =   5033
      TabIndex        =   4
      Top             =   2685
      Width           =   1890
   End
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
      Left            =   2025
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   720
      Width           =   2040
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4725
      OleObjectBlob   =   "VC03.frx":0068
      Top             =   2790
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":029C
      TabIndex        =   5
      Top             =   780
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
      Height          =   225
      Left            =   4140
      OleObjectBlob   =   "VC03.frx":0302
      TabIndex        =   6
      Top             =   660
      Width           =   2955
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":0368
      TabIndex        =   7
      Top             =   2010
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   240
      Left            =   135
      OleObjectBlob   =   "VC03.frx":03D0
      TabIndex        =   8
      Top             =   1470
      Width           =   1695
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
      Height          =   285
      Left            =   75
      OleObjectBlob   =   "VC03.frx":043A
      TabIndex        =   10
      Top             =   45
      Width           =   2325
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
      Height          =   285
      Left            =   4650
      OleObjectBlob   =   "VC03.frx":04B0
      TabIndex        =   11
      Top             =   45
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Height          =   465
      Left            =   -90
      TabIndex        =   12
      Top             =   -45
      Width           =   7350
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
      Left            =   248
      TabIndex        =   3
      Top             =   2655
      Width           =   1890
   End
   Begin VB.PictureBox Picture1 
      Height          =   810
      Left            =   -315
      ScaleHeight     =   750
      ScaleWidth      =   8145
      TabIndex        =   9
      Top             =   2535
      Width           =   8205
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
      Height          =   225
      Left            =   4140
      OleObjectBlob   =   "VC03.frx":0524
      TabIndex        =   15
      Top             =   945
      Width           =   2955
   End
End
Attribute VB_Name = "VC03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lokasi As String
Dim A, Isi, Pusing As String

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
If CCur(Text3) < 0 Then
    MsgBox "NOMINAL BAYAR KURANG", vbCritical, "KONFIRMASI"
    Text3 = Format(CCur(Text1) * CCur(A), "##,###.00")
    Text2.SetFocus
    
    Exit Sub
End If

If Text2 = "0.00" Or Combo1 = "" Or Text1 = "" Or Text2 = "" Then
    MsgBox "MASIH ADA DATA YANG KOSONG", vbCritical, "KONFIRMASI"
    Text1.SetFocus
Exit Sub
End If

Call SimpanVC03
Call EditVC01
Call EditNoBukti

Unload Me
VC03.Show 1
End Sub

Private Sub SimpanVC03()
SSave = "Select * From VC03"
Set RSave = RDCO.OpenResultset(SSave, rdOpenDynamic, rdConcurRowVer)
RSave.AddNew
        RSave("Tanggal") = Date
        RSave("NoNota") = Trim(SkinLabel16)
        RSave("Kode") = Trim(Combo1)
        RSave("Nama") = Trim(SkinLabel3)
        RSave("Jumlah") = CCur(Text1)
        RSave("Jual") = CCur(SkinLabel7)
        RSave("Total") = CCur(Text1) * CCur(SkinLabel7)
RSave.Update
RSave.Close
Set RSave = Nothing
End Sub

Private Sub EditVC01()
Dim Stock As String
Dim HBeli As String

SSave2 = "Select * From VC01 where Kode = '" + Trim(Combo1) + "'"
Set RSave2 = RDCO.OpenResultset(SSave2, rdOpenKeyset, rdConcurRowVer)
    Stock = RSave2("Stokbel")
    HBeli = RSave2("Satuan")
RSave2.Edit
    RSave2("Stokbel") = CCur(Stock) - (CCur(Text1) * CCur(HBeli))
    RSave2("Jumlah") = CCur(Pusing) - CCur(Text1)
RSave2.Update
RSave2.Close
Set RSave2 = Nothing
End Sub

Private Sub EditNoBukti()
SCari9 = "Select * From C013 where Nama = '" + Trim(Operator) + "'"
Set RCari9 = RDCO.OpenResultset(SCari9, rdOpenKeyset, rdConcurRowVer)
    TOGEL = RCari9("NoJual")
    RCari9.Edit
        RCari9("NoJual") = TOGEL + 1
    RCari9.Update
    RCari9.Close
    Set RCari9 = Nothing
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
    SkinLabel7 = Format(RCari2("Jual"), "##,###.00")
    A = RCari2("Jual")
    Isi = RCari2("Jual")
    Pusing = RCari2("Jumlah")
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
Call NoBukti

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
Text3 = "0.00"
SkinLabel6 = ""

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

Private Sub NoBukti()
Dim No As Double
SqlNo = "Select * from C013 where nama = '" + Operator + "'"
Set RSLNO = RDCO.OpenResultset(SqlNo, rdOpenDynamic, rdConcurRowVer)
No = Val(RSLNO("NoJual")) + 1
NoStr = Digit(7, No)
SkinLabel16 = "1." + NoStr
RSLNO.Close
Set RSLNO = Nothing
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
If CCur(Pusing) < CCur(Text1) Then
    MsgBox "JUMLAH DEPOSIT TERSISA " + Pusing + " TRANSAKSI", vbCritical, "KONFIRMASI"
    Text1 = ""
    Text1.SetFocus
    Exit Sub
Else
    If Text1 = "" Then
        Text1 = "0"
    Else
        Text3 = Format(CCur(Text1) * CCur(A), "##,###.00")
        SkinLabel6 = "TOTAL BAYAR :"
    End If
End If
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
If Text2 = "" Then
    Text2 = "0.00"
Else
    Text3 = Format(CCur(Text2) - CCur(Text3), "##,###.00")
    SkinLabel6 = "KEMBALI :"
End If
End Sub
