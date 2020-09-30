Attribute VB_Name = "mdlVariables"
Option Explicit
Public gObjConn As ADODB.Connection
Public gClsConn As New ClsConnection
Public gQuery As String
Public gResult As Clsresultset

'country_master                     - done
Public sCountry(1 To 3) As String
Public sCurrency(1 To 3) As String

'product_master                     - done
Public sProduct(1 To 2) As String

'game_variables                     - done
Public session_id As Integer
Public timestamp As Date

'master_variables                   - done
Public nPrice(1 To 3, 1 To 2) As Double
Public nDemand(1 To 3, 1 To 2) As Double
Public nSupply(1 To 3, 1 To 2) As Double

'shipping_costs_master              - done
'changed by kida to avoid some bugs!! :-)
'Public shippingCost(1 To 3, 1 To 3) as double <--- Original code
Public ShippingCost(0 To 3, 0 To 3) As Double  ' <--- changed code

'cost_master                        - done
Public productionCosts(1 To 12, 1 To 5) As Double
    ' col 1: Labour
    ' col 2: Overtime
    ' col 3: RawMaterial
    ' col 4: Fixed_costs
    ' col 5: Inventory_Costs

'balance_sheet
Public nCash(1 To 12) As Double
Public nInventory(1 To 12) As Double


'latest status of cash
Public nCashCurrent As Double

'form variables
Public bDataEntered As Boolean
Public bCapitalChanged As Boolean


' team_master
Public team_id As Integer
Public team_name As String
Public team_rank As Integer
Public product_id As Double
Public capacity As Double
Public plant_cost As Double
Public country_name(3) As String

Public nSessionID As Integer
Public cash As Double
Public country_id As Integer
Public area_cov(3) As Integer
Public area_cov_cost(3) As Double
Public predicted_price(3) As Double
Public adv(3) As Double
Public promotion(3) As Double
Public csr(3) As Double
Public quality_pu(3) As Double
Public grey_mkt(3) As Double
Public unmet_demand(3) As Double
Public actual_price(3) As Double
Public salesimpact_inventory(3) As Double

Public fixed_cost(20 To 100) As Double
Public manuf_cost_pu As Double
Public OH_cost_pu As Double
Public labour_cost_pu As Double
Public total_manuf_cost As Double
Public total_OH_cost As Double
Public total_labour_cost As Double
Public factoryScrn_Quantity As Double
Public factoryScrn_Quality As Double
Public HumanWelfInit As Double
Public Total_Quant_Ship As Double
Public Quant_Ship(3) As Double
Public chk_insurance(3) As Integer
' this var is used to chk whether a consultant has been used for a particular market
' for a particular session. This wud also be used to calculate the consultancy cost
Public call_consultant(3) As Integer
Public consultancyCost As Double
' this is used to get the info for consultant: for each session we have for
' each market and for the product in the market
Public consultant_info(3) As String
' this is the total shipping cost for the three routes
Public ShipmentCost(3) As Double
Public InsuranceCost(3) As Double

' chk if Long term debt was due for repayment in this session
Public LTrepayment_due As Integer

' chk if Short term debt was due for repayment
Public STrepayment_due As Integer

'short term debt rate for this session and the country
Public shortTermRate As Double

' shipping_costs_master
' this is shipping cost per unit
Public shipping_cost(3) As Double
Public from_country_id As Integer
Public to_country_id As Integer

' news
Public news_headlines(20) As String
Public news_body(20) As String
Public smiley_code(20) As String
Public no_of_news As Integer

'capital
Public nDebt As Double
Public nEquity As Double
Public nSTDebt As Double
Public nRepayment_session_Id As Integer

'Long Term Debt Rate
Public nLTRate(3 To 9) As Single

' upgrade_costs
Public upgrade_cost_per_session As Double
Public upgrade_cost As Double

' engineer's fees
Public engineers_fees As Double

Public newCapacity As Integer

Public RepaymentSessionIndex As Integer             'JUST TO BE USED FOR FINANCE SCREEN TO STORE INDEX OF REPAYMENT SESSION COMBOBOX


Public Function populateArrays(Optional errMsg As String) As Boolean
    Dim nCount As Integer
    
    On Error GoTo ErrHandler
    
    'filling the product master
    gQuery = "select * from product_master"
    Set gResult = Nothing

gClsConn.getDbConn
    Set gResult = gClsConn.executeSQL(gQuery)
    
    If gResult Is Nothing Then
        errMsg = "Please add some entries in product_master"
        GoTo ErrHandler
    Else
        With gResult
            For nCount = 1 To gResult.rowCount
                sProduct(.GetValue(nCount, "product_id")) = .GetValue(nCount, "product_name")
            Next nCount
        End With
    End If

  
    'filling the country master
    gQuery = "select * from country_master"
    Set gResult = Nothing
    
    Set gResult = gClsConn.executeSQL(gQuery)
    
    If gResult Is Nothing Then
        errMsg = "Please add some entries in country_master"
        GoTo ErrHandler
    Else
        With gResult
            For nCount = 1 To gResult.rowCount
                sCountry(.GetValue(nCount, "country_id")) = .GetValue(nCount, "country_name")
                sCurrency(.GetValue(nCount, "country_id")) = .GetValue(nCount, "currency")
            Next nCount
        End With
    End If

gClsConn.closeDbConn
    populateArrays = False
Exit Function

ErrHandler:
    populateArrays = True
End Function

Public Function readDatafromDatabase()

    Dim nCount As Integer
    
  '--------------- game_variables -----------------
    gQuery = "SELECT session_id, timestamp FROM game_variables"
    
    Set gResult = gClsConn.executeSQL(gQuery)
    
    session_id = gResult.GetValue(1, "session_id")
    timestamp = gResult.GetValue(1, "timestamp")
    
    
  '---------------- master_variables --------------
    gQuery = "SELECT s.product_id, s.country_id, s.demand, s.price, s.supply FROM master_variables s" _
        & " WHERE (s.session_id=" & session_id - 1 & ");"
    
    Set gResult = gClsConn.executeSQL(gQuery)
    
    For nCount = 1 To gResult.rowCount
        nPrice(gResult.GetValue(nCount, "country_id"), gResult.GetValue(nCount, "product_id")) _
            = gResult.GetValue(nCount, "price")
        nDemand(gResult.GetValue(nCount, "country_id"), gResult.GetValue(nCount, "product_id")) _
            = gResult.GetValue(nCount, "demand")
        nSupply(gResult.GetValue(nCount, "country_id"), gResult.GetValue(nCount, "product_id")) _
            = gResult.GetValue(nCount, "supply")
    Next nCount

  
  '----------------- shipping_costs_master --------------
    gQuery = "SELECT from_country_id, to_country_id, shipping_cost " _
        & "FROM shipping_costs_master WHERE session_id=" & session_id
    
    Set gResult = gClsConn.executeSQL(gQuery)
    
    For nCount = 1 To gResult.rowCount
        ShippingCost(gResult.GetValue(nCount, "from_country_id"), _
            gResult.GetValue(nCount, "to_country_id")) = gResult.GetValue(nCount, "shipping_cost")
    Next nCount
    
    
  

End Function

Public Function populateObject(objName As Variant, tableName As String, Optional errMsg As String) As Boolean
    Dim nCount As Integer
        
    On Error GoTo ErrHandler
 
Select Case tableName
  Case "country_master"
    With objName
        For nCount = 1 To UBound(sCountry)
            .AddItem sCountry(nCount)
            .ItemData(.NewIndex) = nCount
        Next nCount
    End With
  Case "product_master"
    With objName
        For nCount = 1 To UBound(sProduct)
            .AddItem sProduct(nCount)
            .ItemData(.NewIndex) = nCount
        Next nCount
    End With
End Select

    populateObject = False
    Exit Function
    
ErrHandler:
    populateObject = True
    errMsg = tableName
End Function

Public Function closeSession(errMsg As String)

End Function


Public Function updateCash(Optional errMsg As String)
    
    Dim totalCosts As Variant
    Dim totalCapIncrease As Variant
    Dim oldEquity As Double
    Dim oldDebt As Double
    Dim consultantCost As Double
    
    consultantCost = 20000# * CDbl(call_consultant(0) + call_consultant(1) + call_consultant(2))
    
    'reduce all expenses
    totalCosts = area_cov_cost(1) + area_cov_cost(2) + area_cov_cost(0) + adv(1) + adv(2) + adv(0) + promotion(1) + promotion(2) _
        + promotion(0) + csr(1) + csr(2) + csr(0) + fixed_cost(capacity) + total_labour_cost + total_manuf_cost + total_OH_cost + HumanWelfInit _
        + ShipmentCost(1) + ShipmentCost(2) + ShipmentCost(0) + InsuranceCost(1) + InsuranceCost(2) + InsuranceCost(0) _
        + consultantCost + engineers_fees + upgrade_cost
    
    'increase all cap increase
    gQuery = "SELECT equity, debt FROM balance_sheet WHERE team_id=" & team_id & " session_id=" & nSessionID - 1 & " "
    Set gResult = gClsConn.executeSQL(gQuery)
    
    If gResult Is Nothing Then
    
    Else
        oldEquity = gResult.GetValue(1, "equity")
        oldDebt = gResult.GetValue(1, "debt")
    End If
    
    totalCapIncrease = nEquity - oldEquity + nDebt - oldDebt + nSTDebt
    
    nCashCurrent = cash - totalCosts + totalCapIncrease
    
    frmMain.C_CashStatus = "Xe " & Format(nCashCurrent, "#,##,##0.00")
    
    Select Case nCashCurrent
        Case Is > 0
            frmMain.C_CashStatus.BackColor = &HFF8080
        Case Is <= 0
            frmMain.C_CashStatus.BackColor = vbRed
    End Select
End Function
