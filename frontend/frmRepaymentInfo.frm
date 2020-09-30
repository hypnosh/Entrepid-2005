VERSION 5.00
Begin VB.Form frmRepaymentInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Repayment Information"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepaymentInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   255
      Left            =   7560
      TabIndex        =   1
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
   End
End
Attribute VB_Name = "frmRepaymentInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Dim str As String

If (LTrepayment_due = 1 And STrepayment_due = 1) Then
str = "The Long Term and the Short Term Debt due for payment in this session have been paid back to the bank and deducted from your cash account."
Else
If (LTrepayment_due <> 1 And STrepayment_due = 1) Then
str = "The Short Term Debt due for payment in this session has been paid back to the bank and deducted from your cash account."
Else
If (LTrepayment_due = 1 And STrepayment_due <> 1) Then
str = "The Long Term Debt due for payment in this session has been paid back to the bank deducted from your cash account."
End If
End If
End If
Label1.Caption = str
End Sub

Private Sub Form_Unload(Cancel As Integer)

frmMain.Show

End Sub
