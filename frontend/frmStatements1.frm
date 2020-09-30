VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmStatements 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Finance Module"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7125
   Icon            =   "frmStatements1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbRepaySession 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtSTDebt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox txtDebt 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.TextBox txtEquity 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   720
      TabIndex        =   10
      Top             =   5160
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      BackColor       =   13160660
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "P&&L Statement"
      TabPicture(0)   =   "frmStatements1.frx":120A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "cmdPLGo"
      Tab(0).Control(1)=   "cmbSession"
      Tab(0).Control(2)=   "mshfprofitloss"
      Tab(0).Control(3)=   "Label4"
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Balance Sheet"
      TabPicture(1)   =   "frmStatements1.frx":1226
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "mshfbalancesheet"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmbSessionbs"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdBSgo"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdBSgo 
         Caption         =   "Go"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbSessionbs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdPLGo 
         Caption         =   "Go"
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
         Left            =   -72720
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbSession 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   -73920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshfbalancesheet 
         Height          =   3855
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6800
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshfprofitloss 
         Height          =   3915
         Left            =   -74880
         TabIndex        =   4
         Top             =   960
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   6906
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label1 
         Caption         =   "Session"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   555
      End
      Begin VB.Label Label4 
         Caption         =   "Session"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74760
         TabIndex        =   5
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label lblRepaySession 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BackStyle       =   0  'Transparent
      Caption         =   "Repayment Session"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   5520
      Width           =   1965
   End
   Begin VB.Label lblRate 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0 %"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   6480
      TabIndex        =   17
      Top             =   5520
      Width           =   540
   End
   Begin VB.Label lblSTD 
      BackStyle       =   0  'Transparent
      Caption         =   "ST"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2640
      TabIndex        =   13
      Top             =   5160
      Width           =   2010
   End
   Begin VB.Label lblDebt 
      BackStyle       =   0  'Transparent
      Caption         =   "Debt"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5520
      Width           =   345
   End
   Begin VB.Label lblEquity 
      BackStyle       =   0  'Transparent
      Caption         =   "Equity"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   435
   End
End
Attribute VB_Name = "frmStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Double

Private Sub cmbRepaySession_Click()
    If cmbRepaySession.ListIndex > 0 Then
    
        nRepayment_session_Id = cmbRepaySession.Text
    
        lblRate.Caption = nLTRate(nRepayment_session_Id) & " %"
    Else
        lblRate.Caption = ""
    End If
    
    RepaymentSessionIndex = cmbRepaySession.ListIndex + 1
    
End Sub

Private Sub cmdBSgo_Click()
If cmbSessionbs.ListIndex > -1 Then
    gClsConn.getDbConn
    
     With mshfbalancesheet
     
        gQuery = "SELECT * FROM Balance_Sheet WHERE team_id = " & team_id & " AND session_id = " & cmbSessionbs.Text
            Set gResult = gClsConn.executeSQL(gQuery)
            
        .Col = 1
        .Row = 1
        .Text = Format(gResult.GetValue(1, "cash"), "#,##,##0.00")
        total = gResult.GetValue(1, "cash")
        .Row = 2
        .Text = Format(gResult.GetValue(1, "inventory"), "#,##,##0.00")
        total = total + gResult.GetValue(1, "inventory")
        .Row = 3
        .Text = Format(gResult.GetValue(1, "fixed_assets"), "#,##,##0.00")
        total = total + gResult.GetValue(1, "fixed_assets")
        .Row = 7
        .Text = Format(total, "#,##,##0.00")
        
        .Col = 3
        .Row = 1
        .Text = Format(gResult.GetValue(1, "equity"), "#,##,##0.00")
        .Row = 3
        .Text = Format(gResult.GetValue(1, "previous_retained"), "#,##,##0.00")
        .Row = 4
        .Text = Format(gResult.GetValue(1, "profit_loss"), "#,##,##0.00")
        .Row = 5
        .Text = Format(gResult.GetValue(1, "adjustments"), "#,##,##0.00")
        .Row = 6
        .Text = Format(gResult.GetValue(1, "debt"), "#,##,##0.00")
        .Row = 7
        .Text = Format(total, "#,##,##0.00")
            
     End With
    gClsConn.closeDbConn
End If

End Sub

Private Sub cmdPLGo_Click()
If cmbSession.ListIndex > -1 Then
    gClsConn.getDbConn
        With mshfprofitloss
        gQuery = "SELECT * FROM Profit_Loss t WHERE t.team_id = " & team_id & " AND t.session_id = " _
            & cmbSession.Text & " "
            Set gResult = gClsConn.executeSQL(gQuery)
            
            .Col = 2
            .Row = 1
            .Text = Format(Val(gResult.GetValue(1, "Revenue")) + Val(gResult.GetValue(1, "misc_income")), "#,##,##0.00")
            .Col = 1
            .Row = 3
            .Text = Format(gResult.GetValue(1, "Raw_Material"), "#,##,##0.00")
            .Row = 4
            .Text = Format(gResult.GetValue(1, "Labour_costs"), "#,##,##0.00")
            .Row = 5
            .Text = Format(gResult.GetValue(1, "Shipping_costs"), "#,##,##0.00")
            .Row = 6
            .Text = Format(gResult.GetValue(1, "Market_costs"), "#,##,##0.00")
            .Row = 7
            .Text = Format(gResult.GetValue(1, "Inventory_costs"), "#,##,##0.00")
            .Row = 8
            .Text = Format(gResult.GetValue(1, "Interest_costs"), "#,##,##0.00")
            .Row = 9
            .Text = Format(gResult.GetValue(1, "Other_costs"), "#,##,##0.00")
            .Row = 10
            .Text = Format(gResult.GetValue(1, "tax_costs"), "#,##,##0.00")
            .Row = 11
            .Col = 2
            .Text = Format(gResult.GetValue(1, "Profit_Loss"), "#,##,##0.00")
    
        End With
    gClsConn.closeDbConn
End If
End Sub

Private Sub Form_Load()
    Dim flag As Boolean
    Dim errMsg As String
    Dim i
    
    lblSTD.Caption = "Short Term Debt (" & shortTermRate & " %)"
    For i = 0 To nSessionID - 1
        cmbSession.AddItem i
        cmbSessionbs.AddItem i
    Next i
    
    If nSessionID <= 7 Then
    
        cmbRepaySession.AddItem "-"
        For i = nSessionID + 2 To 9
            cmbRepaySession.AddItem i
        Next i
                'lblRepaySession.Visible = True
        Me.Height = 6210
    Else
        cmbRepaySession.Visible = False
        txtDebt.Visible = False
        udDebt.Visible = False
        
        lblDebt.Visible = False
        lblRepaySession.Visible = False
        Me.Height = 5820
    End If

    With mshfprofitloss
        .Cols = 3
        .Rows = 12
        
        .Col = 1
        .ColWidth(1) = 2000
        .Col = 2
        .ColWidth(2) = 2000
        
        
        .Col = 0
        .ColWidth(0) = 2000
        .Text = "Revenue"
        .Row = 2
        .Text = "Less"
        .Row = 3
        .Text = "   Raw Material"
        .Row = 4
        .Text = "   Labour Costs"
        .Row = 5
        .Text = "   Freight"
        .Row = 6
        .Text = "   Marketing Costs"
        .Row = 7
        .Text = "   Inventory Costs"
        .Row = 8
        .Text = "   Interest Costs"
        .Row = 9
        .Text = "   Other Costs"
        .Row = 10
        .Text = "   Taxes"
        .Row = 11
        .Text = "Net Profit / (Loss)"

    End With
    
    
    With mshfbalancesheet
        .Cols = 4
        .Rows = 8
        
        .Col = 2
        For i = 1 To 7
        .Row = i
        .CellBackColor = &H8000000F
        
        Next i
        
        .ColWidth(0) = 1500
        .ColWidth(1) = 1500
        .ColWidth(2) = 1500
        .ColWidth(3) = 1500
        
        .Row = 0
        .Col = 0
        .Text = "ASSETS"
        .Col = 2
        .Text = "LIABILITIES"
        
        .Row = 1
        .Col = 0
        .Text = "Cash"
        .Col = 2
        .Text = "Equity"
        
        .Row = 2
        .Col = 0
        .Text = "Inventory"
        .Col = 2
        .Text = "Reserves & Surplus"
        
        .Row = 3
        .Col = 0
        .Text = "Fixed Assets"
        .Col = 2
        .Text = "   Retained Earnings"
        
        .Row = 4
        .Col = 2
        .Text = "   Profit / (Loss)"

        .Row = 5
        .Col = 2
        .Text = "   Adjustments"
                
        .Row = 6
        .Col = 2
        .Text = "Debt"

        .Row = 7
        .Text = "Total"
        .Col = 0
        .Text = "Total"
        
    End With
    
    txtEquity = nEquity
    txtDebt = nDebt
    txtSTDebt = nSTDebt
    cmbRepaySession.ListIndex = RepaymentSessionIndex - 1
    lblRate = IIf(nRepayment_session_Id <> 0, (nRepayment_session_Id) & " %", "")
    lblRepaySession.Caption = "Repayment Session "
End Sub



Private Sub Form_Unload(Cancel As Integer)

    If (Val(txtDebt.Text) = 0) Xor (cmbRepaySession.ListIndex < 1) Then
        Cancel = 1
    End If
End Sub


Private Sub txtDebt_Change()
    If txtEquity.Text = "" Then txtEquity.Text = 0
    
    nDebt = Val(txtDebt.Text)
    
    updateCash
End Sub

Private Sub txtDebt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc("."), 8
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtEquity_Change()
    If txtEquity.Text = "" Then txtEquity.Text = 0
    
    nEquity = Val(txtEquity.Text)
    
    updateCash
End Sub

Private Sub txtEquity_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc("."), 8
        
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtSTDebt_Change()
    If txtSTDebt.Text = "" Then txtSTDebt.Text = 0
    
    nSTDebt = Val(txtSTDebt.Text)
    
    updateCash
End Sub

Private Sub txtSTDebt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("0") To Asc("9"), Asc("."), 8
        
        Case Else
            KeyAscii = 0
    End Select
End Sub
