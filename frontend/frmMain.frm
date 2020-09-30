VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Graph"
   ClientHeight    =   11520
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":120A
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdRoute 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoute 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdRoute 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5880
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   25
      Left            =   14640
      Top             =   10680
   End
   Begin VB.CommandButton cmdCommit 
      Caption         =   "Commit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox pctGraphButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   720
      Picture         =   "frmMain.frx":24124C
      ScaleHeight     =   645
      ScaleWidth      =   645
      TabIndex        =   1
      Top             =   0
      Width           =   675
   End
   Begin VB.PictureBox picRouteButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   675
      Left            =   0
      Picture         =   "frmMain.frx":242A76
      ScaleHeight     =   39.49
      ScaleMode       =   0  'User
      ScaleWidth      =   39.49
      TabIndex        =   0
      Tag             =   "0"
      Top             =   0
      Width           =   675
   End
   Begin VB.Label lblCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Country 3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   5880
      TabIndex        =   15
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label lblMarket 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   6720
      TabIndex        =   12
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Image imgMarket 
      Height          =   1395
      Index           =   3
      Left            =   6090
      Picture         =   "frmMain.frx":2442A0
      Top             =   1680
      Width           =   1605
   End
   Begin VB.Label lblCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Country 2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   6960
      TabIndex        =   14
      Top             =   8640
      UseMnemonic     =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblMarket 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   7800
      TabIndex        =   11
      Top             =   9720
      Width           =   1215
   End
   Begin VB.Image imgMarket 
      Height          =   1440
      Index           =   2
      Left            =   7260
      Picture         =   "frmMain.frx":245955
      Top             =   8520
      Width           =   1680
   End
   Begin VB.Label lblCountryName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Country 1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   10200
      TabIndex        =   13
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label lblMarket 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   11280
      TabIndex        =   10
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Image imgMarket 
      Height          =   1365
      Index           =   1
      Left            =   10620
      Picture         =   "frmMain.frx":247109
      Top             =   4635
      Width           =   1605
   End
   Begin VB.Label lblProduction 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   17
      Top             =   9000
      Width           =   1770
   End
   Begin VB.Label cmdNewsTicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3000
      TabIndex        =   16
      Top             =   11280
      Width           =   10335
   End
   Begin VB.Label lblUpgrade 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Upgrade"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4005
      TabIndex        =   9
      Top             =   8760
      Width           =   900
   End
   Begin VB.Label lblFactoryInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   8760
      Width           =   900
   End
   Begin VB.Label C_TeamName 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   11280
      Width           =   3015
   End
   Begin VB.Label C_CashStatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Xe 0.00"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   0
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   13320
      TabIndex        =   6
      Top             =   11280
      Width           =   2055
   End
   Begin VB.Image imgRoute 
      Height          =   7350
      Index           =   2
      Left            =   765
      Picture         =   "frmMain.frx":2488E1
      Top             =   1500
      Visible         =   0   'False
      Width           =   9750
   End
   Begin VB.Image imgRoute 
      Height          =   7230
      Index           =   8
      Left            =   7410
      Picture         =   "frmMain.frx":332163
      Top             =   1485
      Visible         =   0   'False
      Width           =   3645
   End
   Begin VB.Image imgRoute 
      Height          =   6900
      Index           =   3
      Left            =   1680
      Picture         =   "frmMain.frx":3883DD
      Top             =   2820
      Visible         =   0   'False
      Width           =   8970
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1350
      Index           =   1
      Left            =   0
      Picture         =   "frmMain.frx":451F4F
      Top             =   600
      Width           =   1785
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1350
      Index           =   5
      Left            =   1305
      Picture         =   "frmMain.frx":453EA8
      Top             =   2115
      Width           =   1875
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1365
      Index           =   0
      Left            =   13620
      Picture         =   "frmMain.frx":455F09
      Top             =   2970
      Width           =   1740
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1305
      Index           =   4
      Left            =   3075
      Picture         =   "frmMain.frx":457F03
      Top             =   7470
      Width           =   1845
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1335
      Index           =   2
      Left            =   195
      Picture         =   "frmMain.frx":45A11E
      Top             =   8760
      Width           =   1875
   End
   Begin VB.Image imgRoute 
      Height          =   6300
      Index           =   7
      Left            =   7740
      Picture         =   "frmMain.frx":45C1E3
      Top             =   2460
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.Image imgRoute 
      Height          =   6255
      Index           =   1
      Left            =   7680
      Picture         =   "frmMain.frx":4D5215
      Top             =   2385
      Visible         =   0   'False
      Width           =   6465
   End
   Begin VB.Image imgRoute 
      Height          =   6765
      Index           =   4
      Left            =   6870
      Picture         =   "frmMain.frx":559167
      Top             =   2820
      Visible         =   0   'False
      Width           =   4770
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1365
      Index           =   3
      Left            =   9945
      Picture         =   "frmMain.frx":5C25DD
      Top             =   8940
      Width           =   1875
   End
   Begin VB.Image imgRoute 
      Height          =   6405
      Index           =   5
      Left            =   4110
      Picture         =   "frmMain.frx":5C45E2
      Top             =   3090
      Visible         =   0   'False
      Width           =   6450
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1320
      Index           =   6
      Left            =   13485
      Picture         =   "frmMain.frx":64B128
      Top             =   4905
      Width           =   1845
   End
   Begin VB.Image imgFactory 
      Appearance      =   0  'Flat
      Height          =   1320
      Index           =   7
      Left            =   7995
      Picture         =   "frmMain.frx":64D3FC
      Top             =   900
      Width           =   1845
   End
   Begin VB.Image imgRoute 
      Height          =   6345
      Index           =   6
      Left            =   2535
      Picture         =   "frmMain.frx":64F405
      Top             =   2385
      Visible         =   0   'False
      Width           =   7875
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim strQuery As String
Dim clsresult As Clsresultset
Dim nCount As Integer


' Balance_Sheet
Dim equity As Double
Dim debt As Double
Dim previous_retained As Double
Dim fixed_assets As Double
' Dim cash As Double
Dim inventory As Double
Dim adjustments As Double
Dim profit_loss_BS As Double
' profit_loss
Dim revenue As Double
Dim raw_material As Double
Dim labour_costs As Double
Dim overtime_costs As Double
Dim shipping_costs_PnL As Double
Dim inventory_costs As Double
Dim interest_costs As Double
Dim consultancy_costs As Double
' Dim dividend As Double
Dim profit_loss As Double
Dim other_costs As Double
' news - only 3 news per session

Dim all_headlines As String

Dim bRoute As Boolean

Dim nRoute1PositionX(1 To 8) As Integer
Dim nRoute1PositionY(1 To 8) As Integer
Dim nRoute2PositionX(1 To 8) As Integer
Dim nRoute2PositionY(1 To 8) As Integer
Dim nRoute3PositionX(1 To 8) As Integer
Dim nRoute3PositionY(1 To 8) As Integer

Private Function Write_TeamInfo()

' calculations for production_schedule table before update
Dim manufacturingCost As Double
Dim OHCost As Double
Dim labourCost As Double
Dim Total_Qty_Shipped As Double
Dim capacity_update As Double

manufacturingCost = factoryScrn_Quantity * (manuf_cost_pu + 1 * factoryScrn_Quality)
OHCost = factoryScrn_Quantity * OH_cost_pu
labourCost = factoryScrn_Quantity * labour_cost_pu
For i = 1 To 3
    If (call_consultant(i - 1) = 1) Then
        consultancyCost = consultancyCost + 20000
    End If
Next i
consultancyCost = consultancyCost + engineers_fees
' opening connection
    gClsConn.getDbConn
        
        ' insert for production_schedule :: have to update for fixed costs and consultancy costs
        strQuery = "INSERT INTO production_schedule (session_id,team_id,manuf_costs,OH_costs,Labor_costs,quantity,fixed_costs,consultancy_costs) values(" & nSessionID & ", " & team_id & ", " & manufacturingCost & ", " & OHCost & ", " & labourCost & ", " & factoryScrn_Quantity & ", " & fixed_cost(capacity) & ", " & consultancyCost & " ) "
        If Not (gClsConn.Savedata(strQuery)) Then
            Debug.Print strQuery
        End If
        ' insert for sales_impact
        For i = 1 To 3
            Total_Qty_Shipped = Total_Qty_Shipped + Quant_Ship(i - 1)
        Next i
        If (factoryScrn_Quantity - Total_Qty_Shipped > 0) Then
            Quant_Ship(country_id) = Quant_Ship(country_id) + factoryScrn_Quantity - Total_Qty_Shipped
        End If
        For i = 1 To 3
        If (Quant_Ship(i - 1) > 0) Then
            strQuery = "INSERT INTO sales_impact (session_id,team_id,country_id,quantity,promotion,adv,human_cap,csr,quality_pu,area_cov,predicted_price,grey_mkt, coverage_costs) values(" & nSessionID & ", " & team_id & ", " & i & ", " & Quant_Ship(i - 1) - grey_mkt(i - 1) & ", " & promotion(i - 1) & ", " & adv(i - 1) & ", " & HumanWelfInit & ", " & csr(i - 1) & ", " & factoryScrn_Quality & ", " & area_cov(i - 1) & ", " & predicted_price(i - 1) & ", " & grey_mkt(i - 1) & ", " & area_cov_cost(i - 1) & ") "
            If Not (gClsConn.Savedata(strQuery)) Then
                Debug.Print strQuery
            End If
        End If
        Next i
                
        ' insert for shipping transaction
        For i = 1 To 3
            strQuery = "INSERT INTO shipping_transaction (session_id,team_id,to_country_id,quantity,shipping_cost,insured_flag,insurance_cost) values(" & nSessionID & ", " & team_id & ", " & i & ", " & Quant_Ship(i - 1) & ", " & ShipmentCost(i - 1) & ", " & chk_insurance(i - 1) & ", " & InsuranceCost(i - 1) & ") "
            If Not (gClsConn.Savedata(strQuery)) Then
                Debug.Print strQuery
            End If
        Next i
        
                
        ' insert for capital
            strQuery = "INSERT INTO capital (session_id,team_id,debt,equity,stdebt,repayment_session_Id) values(" & nSessionID & ", " & team_id & ", " & nDebt & ", " & nEquity & ", " & nSTDebt & ", " & nRepayment_session_Id & ") "
            If Not (gClsConn.Savedata(strQuery)) Then
                Debug.Print strQuery
            End If
                   
        If (newCapacity < capacity) Then
            capacity_update = capacity
        Else
            capacity_update = newCapacity
        End If
        ' update cash in team_masters
            strQuery = "UPDATE team_master SET cash = " & nCashCurrent & ", capacity = " & capacity_update & ", plant_cost = " & plant_cost & " where team_id = " & team_id & " "
            If Not (gClsConn.Savedata(strQuery)) Then
                Debug.Print strQuery
            End If

  ' closing connection
    gClsConn.closeDbConn
    
End Function

Private Function Get_TeamInfo(nSession As Integer, game_on As Boolean)

If game_on Then
    session_id = nSession
    End If
' opening connection
    gClsConn.getDbConn
    
'team_master
    
    C_TeamName.Caption = " Session " & nSessionID & " :: " & team_name
    
    ' country_master
    strQuery = "select country_name from country_master"
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        country_name(i - 1) = clsresult.GetValue(i, "country_name")
    Next i
    
    
 'costs_master
    strQuery = "select  cm.capacity, manuf_cost_perunit, labour_cost_perunit, OH_cost_perunit from costs_master cm, team_master tm where tm.team_id = " & team_id & " and session_id = " & nSession & " and tm.product_id = cm.product_id and tm.capacity = cm.capacity"
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        manuf_cost_pu = clsresult.GetValue(i, "manuf_cost_perunit")
        OH_cost_pu = clsresult.GetValue(i, "OH_cost_perunit")
        labour_cost_pu = clsresult.GetValue(i, "labour_cost_perunit")
    Next i
    
'shipping_costs_master
    strQuery = "select shipping_cost, from_country_id, to_country_id from shipping_costs_master where session_id = " & nSession & " and from_country_id = " & country_id & " "
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        from_country_id = clsresult.GetValue(i, "from_country_id")
        to_country_id = clsresult.GetValue(i, "to_country_id")
        shipping_cost(to_country_id - 1) = clsresult.GetValue(1, "shipping_cost")
    Next i

'balance_sheet
    strQuery = "select equity, debt, fixed_assets, cash, inventory, previous_retained, profit_loss, adjustments from balance_sheet where team_id = " & team_id & "  and session_id =  " & nSession - 1 & ";"
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
'        equity = clsresult.GetValue(i, "equity")
'        debt = clsresult.GetValue(i, "debt")
'        previous_retained = clsresult.GetValue(i, "previous_retained")
'        fixed_assets = clsresult.GetValue(i, "fixed_assets")
        cash = clsresult.GetValue(i, "cash")
'        inventory = clsresult.GetValue(i, "inventory")
'        adjustments = clsresult.GetValue(i, "adjustments")
'        profit_loss_BS = clsresult.GetValue(i, "profit_loss")
    Next i
    C_CashStatus.Caption = "Xe " & cash
    
    
''profit_loss
'    If (nSession <> 1) Then
'    strQuery = "select  revenue, raw_material, labour_costs, shipping_costs, inventory_costs, interest_costs, profit_loss, other_costs from profit_loss where team_id = " & team_id & " and session_id = " & nSession - 1 & ""
'    Set clsresult = gClsConn.executeSQL(strQuery)
'    For i = 1 To clsresult.rowCount
'        revenue = clsresult.GetValue(i, "revenue")
'        raw_material = clsresult.GetValue(i, "raw_material")
'        labour_costs = clsresult.GetValue(i, "labour_costs")
'    '    overtime_costs = clsresult.GetValue(i, "overtime_costs")
'        shipping_costs = clsresult.GetValue(i, "shipping_costs")
'        inventory_costs = clsresult.GetValue(i, "inventory_costs")
'        consultancy_costs = clsresult.GetValue(i, "consultancy_costs")
'        ' dividend = clsresult.GetValue(i, "dividend")
'        profit_loss = clsresult.GetValue(i, "profit_loss")
'        other_costs = clsresult.GetValue(i, "other_costs")
'    Next i
'    End If
    
'news
    strQuery = "select headlines, body, smiley_code from news where session_id = " & nSession & ""
    Set clsresult = gClsConn.executeSQL(strQuery)
    
    If clsresult Is Nothing Then
    
    Else
    
        no_of_news = clsresult.rowCount
            For i = 1 To clsresult.rowCount
                news_headlines(i) = clsresult.GetValue(i, "headlines")
                news_body(i) = clsresult.GetValue(i, "body")
                smiley_code(i) = clsresult.GetValue(i, "smiley_code")
                all_headlines = all_headlines + news_headlines(i) + "  ||  "
            Next i
            cmdNewsTicker.Caption = Space(350 - Len(news_headlines(1))) & news_headlines(1)
            cmdNewsTicker.Tag = 1
    End If
    
'upgrade_costs
    strQuery = "select upgrade_cost from upgrade_costs where session_id= " & nSessionID & " "
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        upgrade_cost_per_session = clsresult.GetValue(i, "upgrade_cost")
    Next i
    
'sales_impact
    If (nSession <> 1) Then
    strQuery = "select  country_id, area_cov, predicted_price, adv, promotion, csr, quality_pu, grey_mkt, unmet_demand, actual_price, inventory from sales_impact where team_id = " & team_id & " and session_id = " & nSession - 1 & ""
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
    ' we will get country_id from team_master
        'country_id = clsresult.GetValue(i, "country_id")
'        area_cov(country_id - 1) = clsresult.GetValue(i, "area_cov")
'        predicted_price(country_id - 1) = clsresult.GetValue(i, "predicted_price")
'        adv(country_id - 1) = clsresult.GetValue(i, "adv")
'        promotion(country_id - 1) = clsresult.GetValue(i, "promotion")
'        csr(country_id - 1) = clsresult.GetValue(i, "csr")
'        quality_pu(country_id - 1) = clsresult.GetValue(i, "quality_pu")
'        grey_mkt(country_id - 1) = clsresult.GetValue(i, "grey_mkt")
'        unmet_demand(country_id - 1) = clsresult.GetValue(i, "unmet_demand")
'        actual_price(country_id - 1) = clsresult.GetValue(i, "actual_price")
        salesimpact_inventory(country_id - 1) = clsresult.GetValue(i, "inventory")
    Next i
        End If

' LT Rates
    strQuery = "SELECT * FROM LTDebtRate WHERE country_id=" & country_id & " AND session_id=" & nSessionID & " "
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        Dim repay_session As Integer
        repay_session = clsresult.GetValue(i, "repayment_session_id")
        nLTRate(repay_session) = clsresult.GetValue(i, "rate")
    Next i
    
    
' fixed cost
    strQuery = "select fixed_cost, capacity_level from fixed_cost_master"
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        fixed_cost(clsresult.GetValue(i, "capacity_level")) = clsresult.GetValue(i, "fixed_cost")
    Next i
    
    'consultant
    strQuery = "select consultancy from consultant where session_id = " & nSessionID & " and product_id = " & product_id & " "
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        consultant_info(i - 1) = clsresult.GetValue(i, "consultancy")
    Next i
    
    'short term debt interest
    strQuery = "select interest_rate from interestandtaxrates where session_id = " & nSessionID & " and country_id = " & country_id & " "
    Set clsresult = gClsConn.executeSQL(strQuery)
    For i = 1 To clsresult.rowCount
        shortTermRate = clsresult.GetValue(i, "interest_rate")
    Next i
    
    
 ' upgrade
    
    ' closing connection
    gClsConn.closeDbConn
End Function

Private Sub C_CashStatus_Click()
    frmStatements.Show vbModal
End Sub

Private Sub cmdRoute_Click(Index As Integer)
If (factoryScrn_Quantity = 0) Then
    MsgBox "Please manufacture some quantity to ship!"
Else
frmShipmentScreen.nRouteID = Index
frmShipmentScreen.Top = 768
Dim i As Integer
    
    For i = 768 To 5568 - frmShipmentScreen.Height Step -5
        frmShipmentScreen.Top = i
    Next i
    frmShipmentScreen.Show vbModal

End If
End Sub

Private Sub cmdCommit_Click()
    If (nCashCurrent < 0) Then
        MsgBox "You don't have enough cash. Please raise cash."
        frmStatements.Show vbModal
    Else
        cmdCommit.Enabled = False
        s = Write_TeamInfo()
        
        Load frmDataCommitted
        
        frmDataCommitted.Show
        Unload Me
    End If
  
End Sub



Private Sub lblCountryName_Click(Index As Integer)
    If (session_id = 1 Or session_id = 2) Then
        If (Quant_Ship(country_id - 1) + salesimpact_inventory(country_id - 1) = 0) Then
            MsgBox "Please Manufacture some quantity first!"
        Else
        frmMarketScreen.nCountryID = Index
        frmMarketScreen.Show vbModal
        End If
    Else
        If (Quant_Ship(Index - 1) + salesimpact_inventory(Index - 1) = 0) Then
            MsgBox "Please ship some quantity to the market first!"
        Else
        frmMarketScreen.nCountryID = Index
        frmMarketScreen.Show vbModal
        End If
    End If
End Sub

Private Sub lblUpgrade_Click()
    frmUpgrade.Show vbModal
End Sub

Private Sub pctGraphButton_Click()
    frmGraph.Show
End Sub

Private Sub Timer1_Timer()
    Dim str As String
 '   For i = 1 To no_of_news
  '      cmdNewsTicker.Caption = news_headlines(i - 1)
 '   Next i
    Dim k
        k = cmdNewsTicker.Tag
    str = cmdNewsTicker.Caption
    str_len = Len(str)
    str = Mid$(str, 2, Len(str)) + Left(str, 1)
    
    If Left$(str, Len(news_headlines(k))) = Right$(news_headlines(k), 1) & Space$(Len(news_headlines(k)) - 1) Then
        'news number k over... load news number k+1
        
        k = (k + 1) Mod no_of_news
        k = IIf(k = 0, no_of_news, k)
        cmdNewsTicker.Tag = k
        str = Space(350 - Len(news_headlines(k))) & news_headlines(k)
    End If
    cmdNewsTicker.Caption = str

End Sub

Private Sub cmdNewsTicker_Click()
    
    If no_of_news > 0 Then
        frmNews.lblHeadlines = news_headlines(cmdNewsTicker.Tag)
        frmNews.lblNews = country_name(country_id - 1) & ", Session " & nSessionID & ". " _
            & news_body(cmdNewsTicker.Tag)
        
        frmNews.Left = (Me.ScaleWidth * 16 - frmNews.Width) / 2
        frmNews.Top = Me.ScaleHeight * 16 - frmNews.Height - cmdNewsTicker.Height * 16
        frmNews.Show
    End If
    
End Sub


Private Sub Form_Load()
    cmdCommit.Enabled = True
    gClsConn.getDbConn
        gQuery = "SELECT session_id, game_on FROM game_variables"
        
        Set gResult = gClsConn.executeSQL(gQuery)
        
        If gResult Is Nothing Then
            'what the hell???
            
        Else
            If gResult.GetValue(1, "game_on") = False Then
                MsgBox "The game's not on. Wait till the admin tells you to start the game.", vbOKOnly, "Entrepid"
                End
            Else
                nSessionID = gResult.GetValue(1, "session_id")
            End If
        End If
    gClsConn.closeDbConn
            
        
    
    s = Get_TeamInfo(nSessionID, True)
    Dim i As Integer
    For i = 0 To 7
        
        If (i = team_id - 1) Then
            'imgFactory(i).BorderStyle = 1
            imgFactory(i).Enabled = True
            
            lblFactoryInfo.Top = imgFactory(i).Top + 91
            lblFactoryInfo.Left = imgFactory(i).Left
            lblFactoryInfo.Caption = capacity * 50 & " units"
            
            lblUpgrade.Top = lblFactoryInfo.Top
            lblUpgrade.Left = lblFactoryInfo.Left + 60
            
            s = populateArrays()
            
            lblProduction.Top = lblFactoryInfo.Top + 17
            lblProduction.Left = lblFactoryInfo.Left
            lblProduction.Caption = "Production: " & factoryScrn_Quantity & " units"
        Else
            imgFactory(i).Visible = False
        End If
    Next i
    For i = 1 To 3
    ' this if block is to disable the route button and the other two mkts
        If (session_id = 1 Or session_id = 2) And (i <> country_id) Then
            imgMarket(i).Enabled = False
            lblCountryName(i).Enabled = False
            picRouteButton.Enabled = False
        End If
        imgMarket(i).Appearance = 0
        lblCountryName(i).Caption = country_name(i - 1)
        lblMarket(i).Caption = salesimpact_inventory(i - 1)
    Next i
    
        nRoute1PositionX(1) = 850: nRoute1PositionY(1) = 293: nRoute1PositionX(2) = 383: nRoute1PositionY(2) = 229: nRoute1PositionX(3) = 410: nRoute1PositionY(3) = 493: nRoute1PositionX(4) = 740: nRoute1PositionY(4) = 479: nRoute1PositionX(5) = 499: nRoute1PositionY(5) = 445: nRoute1PositionX(6) = 410: nRoute1PositionY(6) = 264: nRoute1PositionX(7) = 856: nRoute1PositionY(7) = 350: nRoute1PositionX(8) = 656: nRoute1PositionY(8) = 224
        nRoute2PositionX(1) = 744: nRoute2PositionY(1) = 429: nRoute2PositionX(2) = 290: nRoute2PositionY(2) = 361: nRoute2PositionX(3) = 296: nRoute2PositionY(3) = 621: nRoute2PositionX(4) = 636: nRoute2PositionY(4) = 614: nRoute2PositionX(5) = 395: nRoute2PositionY(5) = 576: nRoute2PositionX(6) = 334: nRoute2PositionY(6) = 393: nRoute2PositionX(7) = 729: nRoute2PositionY(7) = 491: nRoute2PositionX(8) = 531: nRoute2PositionY(8) = 350
        nRoute3PositionX(1) = 693: nRoute3PositionY(1) = 198: nRoute3PositionX(2) = 244: nRoute3PositionY(2) = 131: nRoute3PositionX(3) = 254: nRoute3PositionY(3) = 378: nRoute3PositionX(4) = 571: nRoute3PositionY(4) = 388: nRoute3PositionX(5) = 353: nRoute3PositionY(5) = 346: nRoute3PositionX(6) = 283: nRoute3PositionY(6) = 168: nRoute3PositionX(7) = 710: nRoute3PositionY(7) = 251: nRoute3PositionX(8) = 514: nRoute3PositionY(8) = 114

    cmdRoute(1).Left = nRoute1PositionX(team_id)
    cmdRoute(1).Top = nRoute1PositionY(team_id)
    cmdRoute(2).Left = nRoute2PositionX(team_id)
    cmdRoute(2).Top = nRoute2PositionY(team_id)
    cmdRoute(3).Left = nRoute3PositionX(team_id)
    cmdRoute(3).Top = nRoute3PositionY(team_id)
    
    
    
    'load country & product info
    Dim errMsg As String
    If populateArrays(errMsg) Then
        MsgBox errMsg
    End If
    
    If no_of_news = 0 Then Timer1.Enabled = False
    
    updateCash
End Sub

Private Sub imgFactory_Click(Index As Integer)
    frmFactoryScreen.nFactoryID = Index
    frmFactoryScreen.Show vbModal
End Sub

Private Sub imgMarket_Click(Index As Integer)
    
    If (session_id = 1 Or session_id = 2) Then
        If (Quant_Ship(country_id - 1) + salesimpact_inventory(country_id - 1) = 0) Then
            MsgBox "Please manufacture some quantity first!"
        Else
            frmMarketScreen.nCountryID = Index
            frmMarketScreen.Show vbModal
        End If
    Else
        If (Quant_Ship(Index - 1) + salesimpact_inventory(Index - 1) = 0) Then
            MsgBox "Please ship some quantity to the market first!"
        Else
            frmMarketScreen.nCountryID = Index
            frmMarketScreen.Show vbModal
    End If
    End If
End Sub

Private Sub pctExitButton_Click()
    
    End
End Sub

Private Sub picRouteButton_Click()
    bRoute = Not bRoute
    If bRoute Then
        picRouteButton.Appearance = 1
    Else
        picRouteButton.Appearance = 0
    End If
    
    Dim i As Integer
    For i = 1 To 3
        cmdRoute(i).Visible = bRoute
    Next i
            imgRoute(team_id).Visible = bRoute
End Sub

