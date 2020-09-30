VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmGraph 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Graph Console"
   ClientHeight    =   7215
   ClientLeft      =   30
   ClientTop       =   375
   ClientWidth     =   12270
   Icon            =   "frmGraph.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar Toolbar2 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   12270
      _ExtentX        =   21643
      _ExtentY        =   1058
      ButtonWidth     =   794
      ButtonHeight    =   953
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Bar"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Line"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Area"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Step"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   2
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "3D"
            Key             =   ""
            Object.Tag             =   ""
            Style           =   1
         EndProperty
      EndProperty
   End
   Begin MSChart20Lib.MSChart chtDSP 
      Height          =   6135
      Left            =   0
      OleObjectBlob   =   "frmGraph.frx":120A
      TabIndex        =   4
      Top             =   600
      Width           =   12255
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go!"
      Default         =   -1  'True
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
      Left            =   11520
      TabIndex        =   3
      Top             =   6840
      Width           =   612
   End
   Begin VB.ComboBox cmbType 
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
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   6840
      Width           =   3495
   End
   Begin VB.ComboBox cmbProduct 
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
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   6840
      Width           =   3975
   End
   Begin VB.ComboBox cmbCountry 
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
      ItemData        =   "frmGraph.frx":3715
      Left            =   0
      List            =   "frmGraph.frx":3717
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   6840
      Width           =   3735
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5415
      Left            =   8160
      TabIndex        =   6
      Top             =   720
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmGraph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdGo_Click()
    Dim rResult As Recordset, sWantedValues As String
    
    If cmbCountry.ListIndex = -1 Or cmbProduct.ListIndex = -1 Or cmbType.ListIndex = -1 Then
        'nothing is selected
        Exit Sub
    End If
    
    Select Case cmbType.ListIndex
        Case 0
            'Demand & Supply
            sWantedValues = "demand, supply"
        Case 1
            'Price
            sWantedValues = "price"
    End Select
    
    gQuery = "SELECT " & sWantedValues & " FROM master_variables s" _
        & " WHERE (s.product_id=" & cmbProduct.ItemData(cmbProduct.ListIndex) _
        & ") AND (s.country_id=" & cmbCountry.ItemData(cmbCountry.ListIndex) _
        & ") AND (s.session_id<" & session_id & ")"
    'gQuery = "SELECT s.session_id, s." & LCase(cmbType.Text) & " FROM master_variables s" _
    '    & " WHERE (s.product_id=" & cmbProduct.ItemData(cmbProduct.ListIndex) _
    '    & ") AND (s.country_id=" & cmbCountry.ItemData(cmbCountry.ListIndex) _
    '    & ") AND (s.session_id<" & session_id & ");"
gClsConn.getDbConn
    
    Set gResult = gClsConn.executeSQL(gQuery)   'executes the SQL query
        If gResult Is Nothing Then
            MsgBox "Fill the tables dumbo!"
            End
        End If
        
    Set rResult = gClsConn.recSet   'extracts the recordset of the above query

    Set chtDSP.DataSource = rResult 'sets the chart datasource to the recordset
    Set MSHFlexGrid1.DataSource = rResult
gClsConn.closeDbConn
    chtDSP.Title = "Trends of " & cmbType.Text & " for " & cmbProduct.Text & " in " & cmbCountry.Text
    Me.Caption = "Graph Console: " & chtDSP.Title
End Sub

Private Sub Form_Load()

Dim errMsg As String
Dim flag As Boolean

flag = populateObject(cmbCountry, "country_master", errMsg)

flag = populateObject(cmbProduct, "product_master", errMsg)

    
With cmbType
    .AddItem ("Demand & Supply")
    .AddItem ("Price")
End With
End Sub

Private Sub Toolbar2_ButtonClick(ByVal Button As ComctlLib.Button)
    chtDSP.chartType = Toolbar2.Buttons(1).Value * 0 + Toolbar2.Buttons(2).Value * 2 + Toolbar2.Buttons(3).Value * 4 + Toolbar2.Buttons(4).Value * 6 + (1 - Toolbar2.Buttons(6).Value)
End Sub
