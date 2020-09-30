VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmFactoryScreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Factory Screen"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFactoryScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Slider S_HWI 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   1080
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100000
      TickStyle       =   3
   End
   Begin ComctlLib.Slider S_Quantity 
      Height          =   375
      Left            =   2040
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin ComctlLib.Slider S_Quality 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   600
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin VB.Label Label_HWI 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   7680
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label L_HWI 
      BackStyle       =   0  'Transparent
      Caption         =   "Human Welfare Initiatives"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Fixed Cost"
      Height          =   255
      Left            =   5520
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
   Begin VB.Label L_FixedCost 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5520
      TabIndex        =   12
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Overhead Cost"
      Height          =   255
      Left            =   3720
      TabIndex        =   11
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Labour Cost"
      Height          =   255
      Left            =   1920
      TabIndex        =   10
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Manufacturing Cost"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label L_OC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3720
      TabIndex        =   8
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label L_LC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label L_MC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label_Qual 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7680
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label_Quant 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   255
      Left            =   7680
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label L_Quality 
      BackStyle       =   0  'Transparent
      Caption         =   "Quality"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   570
   End
   Begin VB.Label L_Quantity 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity to be produced"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   1830
   End
End
Attribute VB_Name = "frmFactoryScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nFactoryID As Integer

Private Sub Form_Load()
    L_FixedCost.Caption = "Xe " & Format(fixed_cost(capacity), "#,##,##0.00")
    Me.Caption = "Factory Screen"
    S_Quantity.Max = capacity * 50
    S_Quantity.Value = factoryScrn_Quantity
    S_Quality.Value = factoryScrn_Quality
    S_HWI.Value = HumanWelfInit
    Label_Quant.Caption = S_Quantity & " units"
    Label_Qual.Caption = S_Quality & " %"
    L_MC.Caption = "Xe " & Format(S_Quantity * (manuf_cost_pu + 1 * S_Quality.Value), "#,##,##0.00")
    L_OC.Caption = "Xe " & Format(S_Quantity * OH_cost_pu, "#,##,##0.00")
    L_LC.Caption = "Xe " & Format(S_Quantity * labour_cost_pu, "#,##,##0.00")
    Label_HWI.Caption = "Xe " & Format(S_HWI, "#,##,##0.00")
End Sub

Private Sub Form_Unload(Cancel As Integer)
factoryScrn_Quantity = S_Quantity.Value
factoryScrn_Quality = S_Quality.Value
total_manuf_cost = S_Quantity.Value * (manuf_cost_pu + 1 * S_Quality.Value)
total_OH_cost = S_Quantity.Value * OH_cost_pu
total_labour_cost = S_Quantity * labour_cost_pu
HumanWelfInit = S_HWI.Value
' this code allocates the total productino to be shipped to the home mkt during 1st two sessions
If (session_id = 1 Or session_id = 2) Then
    For i = 1 To 3
    If (i = country_id) Then
        Quant_Ship(i - 1) = factoryScrn_Quantity
    Else
        Quant_Ship(i - 1) = 0
    End If
        chk_insurance(i - 1) = 0
        ShipmentCost(i - 1) = 0
        InsuranceCost(i - 1) = 0
    Next i
End If
updateCash
End Sub


Private Sub S_HWI_Scroll()
    Label_HWI.Caption = "Xe " & Format(S_HWI, "#,##,##0.00")
End Sub


Private Sub S_Quality_Scroll()
    Label_Qual.Caption = Format(S_Quality, "# ") & "%"
    L_MC.Caption = "Xe " & Format(S_Quantity * (manuf_cost_pu + 1 * S_Quality.Value), "#,##,##0.00")
End Sub

Private Sub S_Quantity_Scroll()
    Dim total_Shipment As Double
    Dim i As Integer
    For i = 1 To 3
        total_Shipment = Quant_Ship(i - 1) + total_Shipment
    Next i
    If (S_Quantity.Value < total_Shipment) Then
        S_Quantity.Value = total_Shipment
    End If
    Label_Quant.Caption = S_Quantity & " units"
    ' wt used for quality is 1
    L_MC.Caption = "Xe " & Format(S_Quantity * (manuf_cost_pu + 1 * S_Quality.Value), "#,##,##0.00")
    L_OC.Caption = "Xe " & Format(S_Quantity * OH_cost_pu, "#,##,##0.00")
    L_LC.Caption = "Xe " & Format(S_Quantity * labour_cost_pu, "#,##,##0.00")
    frmMain.lblProduction.Caption = "Production: " & S_Quantity & " units"
End Sub
