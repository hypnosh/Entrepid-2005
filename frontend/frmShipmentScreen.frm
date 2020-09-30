VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmShipmentScreen 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Shipment Screen"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6645
   Icon            =   "frmShipmentScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox C_InsurShipment 
      Caption         =   "Insurance"
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
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   1695
   End
   Begin ComctlLib.Slider S_QtyShip 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   661
      _Version        =   327682
      Max             =   100
      TickStyle       =   3
      TickFrequency   =   2
   End
   Begin VB.Label L_ShipmentCost 
      BackStyle       =   0  'Transparent
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
      Left            =   4200
      TabIndex        =   5
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label L_InsurCost 
      BackStyle       =   0  'Transparent
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label_Quant 
      BackStyle       =   0  'Transparent
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
      Left            =   5760
      TabIndex        =   3
      Top             =   600
      Width           =   855
   End
   Begin VB.Label L_QtyShip 
      BackStyle       =   0  'Transparent
      Caption         =   "Quantity to be Shipped"
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
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2070
   End
End
Attribute VB_Name = "frmShipmentScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nRouteID As Integer
Dim Quant_Allowed As Double

Private Sub C_InsurShipment_Click()
If (C_InsurShipment.Value = 1) Then
    L_InsurCost.Caption = "Insurance Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1) / 2, "#,##,##0.00")
Else
    L_InsurCost.Caption = "Insurance Cost: Xe 0.00"
End If
End Sub

Private Sub Form_Load()
Dim Quant_Rem_Two As Double
S_QtyShip.Max = factoryScrn_Quantity
S_QtyShip.Value = Quant_Ship(nRouteID - 1)
 L_InsurCost.Caption = "Insurance Cost: Xe " & Format(InsuranceCost(nRouteID - 1), "#,##,##0.00")
 L_ShipmentCost.Caption = "Shipment Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1), "#,##,##0.00")
 Label_Quant.Caption = S_QtyShip & " units"
For i = 1 To 3
    If (i <> nRouteID) Then
        Quant_Rem_Two = Quant_Rem_Two + Quant_Ship(i - 1)
    End If
    Quant_Allowed = factoryScrn_Quantity - Quant_Rem_Two
Next i
C_InsurShipment.Value = chk_insurance(nRouteID - 1)
Me.Caption = "Shipment Screen - " & country_name(nRouteID - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quant_Ship(nRouteID - 1) = S_QtyShip.Value
chk_insurance(nRouteID - 1) = C_InsurShipment.Value
ShipmentCost(nRouteID - 1) = S_QtyShip * shipping_cost(nRouteID - 1)
If (C_InsurShipment.Value = 1) Then
    InsuranceCost(nRouteID - 1) = ShipmentCost(nRouteID - 1) / 2
Else
    InsuranceCost(nRouteID - 1) = 0
End If
updateCash
End Sub


Private Sub S_QtyShip_Click()
Label_Quant.Caption = S_QtyShip & " units"
L_ShipmentCost.Caption = "Shipment Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1), "#,##,##0.00")
If (C_InsurShipment.Value = 1) Then
    L_InsurCost.Caption = "Insurance Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1) / 2, "#,##,##0.00")
Else
    L_InsurCost.Caption = "Insurance Cost: Xe 0.00"
End If
If (S_QtyShip.Value > Quant_Allowed) Then
    'MsgBox "Please manufacture more to ship more than " & Quant_Allowed
    S_QtyShip.Value = Quant_Allowed
End If
End Sub

Private Sub S_QtyShip_Scroll()
Label_Quant.Caption = S_QtyShip & " units"
L_ShipmentCost.Caption = "Shipment Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1), "#,##,##0.00")
If (C_InsurShipment.Value = 1) Then
    L_InsurCost.Caption = "Insurance Cost: Xe " & Format(S_QtyShip * shipping_cost(nRouteID - 1) / 2, "#,##,##0.00")
Else
    L_InsurCost.Caption = "Insurance Cost: Xe 0.00"
End If
If (S_QtyShip.Value > Quant_Allowed) Then
    'MsgBox "Please manufacture more to ship more than " & Quant_Allowed
    S_QtyShip.Value = Quant_Allowed
End If
End Sub
