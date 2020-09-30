VERSION 5.00
Begin VB.Form frmTeam 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Team Entrepid"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTeam.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   3855
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lblTeam 
      BackColor       =   &H00000000&
      Caption         =   "The team behind Entrepid"
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
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
   End
End
Attribute VB_Name = "frmTeam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With Label1
        .Caption = "Chakradhar Gade" & vbCrLf & "Dinesh Chaudhary" & vbCrLf & _
            "Manan Gupta" & vbCrLf & "Shivam Arren"
        .Caption = .Caption & vbCrLf & "Adarsh Mohta" & vbCrLf & "Angshuman Goswami" _
            & vbCrLf & "Chirag Surana" & vbCrLf & "Kiran Nanduri" & vbCrLf & _
            "Manish Gupta" & vbCrLf & "Supriya Sahai" & vbCrLf _
            & "Prasad Ajinkya" & vbCrLf & "Amit Sharma"
        .Caption = .Caption & vbCrLf & vbCrLf & "Original Concept: Abhishek Dangra"
    End With
End Sub
