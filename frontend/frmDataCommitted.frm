VERSION 5.00
Begin VB.Form frmDataCommitted 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7725
   Icon            =   "frmDataCommitted.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOpenSession 
      Caption         =   "Open Session"
      Height          =   375
      Left            =   6360
      TabIndex        =   1
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Data has been entered. Wait till the next session is announced open."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1515
      TabIndex        =   0
      Top             =   938
      Width           =   4695
   End
End
Attribute VB_Name = "frmDataCommitted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOpenSession_Click()
    For i = 1 To 3
        call_consultant(i - 1) = 0
        Quant_Ship(i - 1) = 0
        chk_insurance(i - 1) = 0
        InsuranceCost(i - 1) = 0
        ShipmentCost(i - 1) = 0
        adv(i - 1) = 0
        csr(i - 1) = 0
        promotion(i - 1) = 0
        area_cov(i - 1) = 0
        predicted_price(i - 1) = 0
        grey_mkt(i - 1) = 0
        area_cov_cost(i - 1) = 0
    Next i
    factoryScrn_Quality = 0
    factoryScrn_Quantity = 0
    total_labour_cost = 0
    total_manuf_cost = 0
    total_OH_cost = 0
    HumanWelfInit = 0
    nDebt = 0
    nEquity = 0
    nSTDebt = 0
    nRepayment_session_Id = 0
    
    gClsConn.getDbConn
    
      gQuery = "SELECT * FROM team_master WHERE name = '" & frmLogin.txtUserName & "' AND [password] = '" & frmLogin.txtPassword & "' "
        
        Set gResult = Nothing
        Set gResult = gClsConn.executeSQL(gQuery)

        team_rank = gResult.GetValue(1, "rank")
        capacity = gResult.GetValue(1, "capacity")
        plant_cost = gResult.GetValue(1, "plant_cost")
        
        
      gQuery = "SELECT session_id, game_on FROM game_variables"
        
        Set gResult = Nothing
        Set gResult = gClsConn.executeSQL(gQuery)
        
        If gResult Is Nothing Then
            'what the hell???
            
        Else
            If (Not gResult.GetValue(1, "game_on")) Or nSessionID = gResult.GetValue(1, "session_id") Then
                MsgBox "The game's not on. Wait till the admin tells you to start the game.", vbOKOnly, "Entrepid"
                Exit Sub
            Else
                nSessionID = gResult.GetValue(1, "session_id")
                ' Load frmMain
                ' Unload Me
            End If
        End If
        
        ' chk long term repayment session
    gQuery = "select * from capital where team_id = " & team_id & " and repayment_session_Id = " & nSessionID & " "
    Set gResult = gClsConn.executeSQL(gQuery)
    If gResult Is Nothing Then
     LTrepayment_due = 0
    Else
     LTrepayment_due = 1
    End If
    
    ' chk short term repayment session
    gQuery = "select STDebt from capital where team_id = " & team_id & " and session_id = " & nSessionID - 1 & " "
    Set gResult = gClsConn.executeSQL(gQuery)
    If gResult Is Nothing Then
    Else
    If (gResult.GetValue(1, "STDebt") <> 0) Then
        STrepayment_due = 1
    Else
        STrepayment_due = 0
    End If
    End If
    
    
    If (LTrepayment_due <> 1 And STrepayment_due <> 1) Then
    frmMain.Show
    Unload Me
    Else
    frmRepaymentInfo.Show
    Unload Me
    End If

    gClsConn.closeDbConn
End Sub

Private Sub Form_Load()

End Sub
