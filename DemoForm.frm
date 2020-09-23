VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo Form "
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Change Font Properties"
      Height          =   630
      Left            =   1230
      TabIndex        =   1
      Top             =   2355
      Width           =   1725
   End
   Begin VB.TextBox Text1 
      Height          =   1995
      Left            =   465
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "DemoForm.frx":0000
      Top             =   195
      Width           =   3675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
frmFont.Show
End Sub

Private Sub Form_Load()
' Go to the cmdOkay_Click and change to suit your needs.
End Sub
