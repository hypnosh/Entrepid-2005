VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMarketScreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Market Screen"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9510
   Icon            =   "frmMarketScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   9510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider S_PredPrice 
      Height          =   375
      Left            =   2040
      TabIndex        =   18
      Top             =   2040
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Max             =   8000
      TickStyle       =   3
   End
   Begin ComctlLib.Slider S_AreaCov 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Min             =   10
      Max             =   100
      SelStart        =   10
      TickStyle       =   3
      Value           =   10
   End
   Begin VB.CheckBox C_GM 
      Caption         =   "Grey Market"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1455
   End
   Begin ComctlLib.Slider S_Adv 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   120
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Max             =   200000
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin VB.CommandButton C_Consult 
      BackColor       =   &H80000007&
      Caption         =   "Call Consultant"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      MaskColor       =   &H00FF0000&
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
   Begin ComctlLib.Slider S_CSR 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   600
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Max             =   200000
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin ComctlLib.Slider S_Promo 
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1080
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Max             =   200000
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin ComctlLib.Slider S_GM 
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   327682
      Max             =   5000
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin VB.Label Label_PredPrice 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   16
      Top             =   2040
      Width           =   1170
   End
   Begin VB.Label L_PredPrice 
      BackStyle       =   0  'Transparent
      Caption         =   "Predicted Price (per unit)"
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   2160
      Width           =   1815
   End
   Begin VB.Label Label_AreaCov 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   14
      Top             =   1560
      Width           =   1170
   End
   Begin VB.Label L_Area_Cov 
      BackStyle       =   0  'Transparent
      Caption         =   "Area"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label_GM 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   12
      Top             =   2520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label_Promo 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   11
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label Label_CSR 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   10
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label Label_Adv 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   8280
      TabIndex        =   9
      Top             =   120
      Width           =   1170
   End
   Begin VB.Label L_Promotions 
      BackStyle       =   0  'Transparent
      Caption         =   "Promotions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label L_CSR 
      BackStyle       =   0  'Transparent
      Caption         =   "CSR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   465
   End
   Begin VB.Label L_Adv 
      BackStyle       =   0  'Transparent
      Caption         =   "Advertisement"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmMarketScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public chk_GM As Integer
Public nCountryID As Integer
Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub area_cov_Click()

End Sub

Private Sub C_Consult_Click()
    frmConsultant.nCountryID = nCountryID
    frmConsultant.Top = 768
    Dim i As Integer
    For i = 768 To 5568 - frmConsultant.Height Step -5
        frmConsultant.Top = i
    Next i
    frmConsultant.Show vbModal
    If (call_consultant(nCountryID - 1) = 1) Then
        C_Consult.Enabled = False
    End If
End Sub

Private Sub C_GM_Click()

If (C_GM.Value = 1) Then
S_GM.Visible = True
Label_GM.Visible = True
Else
S_GM.Visible = False
Label_GM.Visible = False
End If
End Sub

Private Sub Form_Load()
    If (area_cov(nCountryID - 1) < 10) Then
        area_cov(nCountryID - 1) = 10
    End If
    If (call_consultant(nCountryID - 1) = 1) Then
        C_Consult.Enabled = False
    End If
    Me.Caption = "Market Screen - " & country_name(nCountryID - 1)
    S_Adv.Value = adv(nCountryID - 1)
    S_CSR.Value = csr(nCountryID - 1)
    S_Promo.Value = promotion(nCountryID - 1)
    S_AreaCov.Value = area_cov(nCountryID - 1)
    S_PredPrice.Value = predicted_price(nCountryID - 1)
    S_GM.Value = grey_mkt(nCountryID - 1)
     
    S_GM.Max = Quant_Ship(nCountryID - 1) + salesimpact_inventory(nCountryID - 1)
    If (grey_mkt(nCountryID - 1) <> 0) Then
        chk_GM = 1
    Else
        chk_GM = 0
    End If
    C_GM.Value = chk_GM
    L_Area_Cov.Caption = "Area Covered: " & S_AreaCov & " %"
    Label_Adv.Caption = "Xe " & Format(S_Adv, "#,##,##0.00")
    Label_CSR.Caption = "Xe " & Format(S_CSR, "#,##,##0.00")
    Label_Promo.Caption = "Xe " & Format(S_Promo, "#,##,##0.00")
    Label_AreaCov.Caption = "Xe " & Format((S_AreaCov * 2000), "#,##,##0.00")
    Label_PredPrice.Caption = "Xe " & Format(S_PredPrice, "#,##,##0.00")
    Label_GM.Caption = S_GM & " units"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    adv(nCountryID - 1) = S_Adv.Value
    csr(nCountryID - 1) = S_CSR.Value
    promotion(nCountryID - 1) = S_Promo.Value
    area_cov(nCountryID - 1) = S_AreaCov.Value
    predicted_price(nCountryID - 1) = S_PredPrice.Value
    grey_mkt(nCountryID - 1) = S_GM.Value
    area_cov_cost(nCountryID - 1) = S_AreaCov.Value * 2000
    chk_GM = C_GM.Value
    updateCash
End Sub

Private Sub S_Adv_Scroll()
Label_Adv.Caption = "Xe " & Format(S_Adv, "#,##,##0.00")
End Sub
Private Sub S_CSR_Scroll()
Label_CSR.Caption = "Xe " & Format(S_CSR, "#,##,##0.00")
End Sub
Private Sub S_Promo_Scroll()
Label_Promo.Caption = "Xe " & Format(S_Promo, "#,##,##0.00")
End Sub
Private Sub S_AreaCov_Scroll()
Label_AreaCov.Caption = "Xe " & Format(S_AreaCov * 2000, "#,##,##0.00")
L_Area_Cov.Caption = "Area Covered: " & S_AreaCov & " %"
End Sub
Private Sub S_PredPrice_Scroll()
Label_PredPrice.Caption = "Xe " & Format(S_PredPrice, "#,##,##0.00")
End Sub
Private Sub S_GM_Scroll()
Label_GM.Caption = S_GM & " units"
End Sub

