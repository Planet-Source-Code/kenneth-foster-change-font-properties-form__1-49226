VERSION 5.00
Begin VB.Form frmFont 
   Caption         =   "Change Font Properties"
   ClientHeight    =   5145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   ScaleHeight     =   5145
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   690
      Left            =   5895
      TabIndex        =   53
      Top             =   1920
      Width           =   1590
   End
   Begin VB.ComboBox cmbFonts 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3255
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   52
      Top             =   30
      Width           =   2670
   End
   Begin VB.ComboBox cmbFontSize 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6045
      Style           =   2  'Dropdown List
      TabIndex        =   51
      Top             =   30
      Width           =   675
   End
   Begin VB.CheckBox Bold1 
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3390
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   2475
      Width           =   510
   End
   Begin VB.CheckBox Italic2 
      Caption         =   "I"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3990
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   2475
      Width           =   510
   End
   Begin VB.CheckBox Underline3 
      Caption         =   "U"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4605
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   2475
      Width           =   510
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   765
      Left            =   5895
      TabIndex        =   47
      Top             =   870
      Width           =   1590
   End
   Begin VB.PictureBox CurColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   135
      ScaleHeight     =   315
      ScaleWidth      =   2790
      TabIndex        =   41
      Top             =   1950
      Width           =   2820
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   40
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   39
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   38
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   37
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   5
      Left            =   150
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   36
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   510
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   35
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   6
      Left            =   510
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   34
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   7
      Left            =   510
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   33
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000000C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   8
      Left            =   510
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   32
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   9
      Left            =   510
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   10
      Left            =   870
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   11
      Left            =   870
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   29
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   12
      Left            =   870
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H000040C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   13
      Left            =   870
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00004080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   14
      Left            =   870
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   15
      Left            =   1230
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   16
      Left            =   1230
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   24
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   17
      Left            =   1230
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   18
      Left            =   1230
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   19
      Left            =   1230
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFC0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   20
      Left            =   1590
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   21
      Left            =   1590
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   22
      Left            =   1590
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   23
      Left            =   1590
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00008000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   24
      Left            =   1590
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   25
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF80&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   26
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   27
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   28
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   29
      Left            =   1950
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   30
      Left            =   2310
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF8080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   31
      Left            =   2310
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   32
      Left            =   2310
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   33
      Left            =   2310
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   34
      Left            =   2310
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   105
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   35
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1605
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF80FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   36
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   1230
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF00FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   37
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   855
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00C000C0&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   38
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   480
      Width           =   255
   End
   Begin VB.PictureBox MiniPicker 
      Appearance      =   0  'Flat
      BackColor       =   &H00800080&
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   39
      Left            =   2670
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   105
      Width           =   255
   End
   Begin VB.TextBox HEXvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   150
      MaxLength       =   6
      TabIndex        =   0
      Text            =   "FFFFFF"
      Top             =   2610
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox G 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1065
      MaxLength       =   3
      TabIndex        =   43
      Text            =   "255"
      Top             =   2625
      Width           =   970
   End
   Begin VB.TextBox B 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   255
      Left            =   1950
      MaxLength       =   3
      TabIndex        =   44
      Text            =   "255"
      Top             =   2640
      Width           =   920
   End
   Begin VB.TextBox R 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   270
      MaxLength       =   3
      TabIndex        =   45
      Text            =   "255"
      Top             =   2655
      Width           =   920
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   750
      Left            =   3240
      Shape           =   4  'Rounded Rectangle
      Top             =   2355
      Width           =   2025
   End
   Begin VB.Label lblSample 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sample Text"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   3255
      TabIndex        =   46
      Top             =   3450
      Width           =   1035
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1680
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   3210
      Width           =   7740
   End
   Begin VB.Label Label1 
      Caption         =   "HEX"
      Height          =   225
      Left            =   1305
      TabIndex        =   42
      Top             =   2385
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3060
      Left            =   0
      Top             =   45
      Width           =   3090
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Color Picker code is borrowed from Huovinen Harri aka IONIC

Private Sub cmdOkay_Click()
    Change_Font Form1.Text1   'Change to suit your program
                              'e.g.  Change_Font Form1.Label1
    
   Unload frmFont
End Sub
Public Function GetHEXValue()
    
    Dim HEXr As String, HEXg As String, HEXb As String
    
    HEXr = Hex$(R.Text)
    If Len(HEXr) = 1 Then HEXr = "0" & HEXr
    
    HEXg = Hex$(G.Text)
    If Len(HEXg) = 1 Then HEXg = "0" & HEXg
    
    HEXb = Hex$(B.Text)
    If Len(HEXb) = 1 Then HEXb = "0" & HEXb
    
    HEXvalue.Text = HEXr & HEXg & HEXb
    
End Function
Public Function GetRGBValue()
    
    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = CurColor.BackColor And 255
    ColorG = (CurColor.BackColor And 65280) / 256
    ColorB = (CurColor.BackColor And 16711680) / 65535
    
    R.Text = ColorR
    G.Text = ColorG
    B.Text = ColorB
    
End Function
Private Sub B_Change()
    
    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    GetHEXValue
    
End Sub

Private Sub cmdCancel_Click()
Unload frmFont
End Sub
Private Sub Form_Load()
    Dim I As Integer
    Dim fs As Integer
    
    'font list
    For I = 1 To Screen.FontCount
        cmbFonts.AddItem Screen.Fonts(I - 1)
    Next I
    
    'font size
    For fs = 8 To 50 Step 2   'show only even sizes
        cmbFontSize.AddItem fs
    Next fs
    
    cmbFonts.Text = lblSample.FontName
    cmbFontSize.Text = Int(lblSample.FontSize)  'no fractions
    Underline3.FontUnderline = True  'underlines the U in the caption
    
    HEXvalue.Visible = True
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    Unload frmFont
    Unload Me
End Sub
Private Sub G_Change()
    
    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    GetHEXValue
    
End Sub
Private Sub MiniPicker_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    CurColor.BackColor = MiniPicker(Index).BackColor
    GetRGBValue
    GetHEXValue
    lblSample.ForeColor = CurColor.BackColor
End Sub
Private Sub R_Change()
    
    Dim ColorR As String, ColorG As String, ColorB As String
    
    ColorR = R.Text
    ColorG = G.Text
    ColorB = B.Text
    
    CurColor.BackColor = RGB(ColorR, ColorG, ColorB)
    GetHEXValue
    
End Sub
Private Sub cmbFonts_Click()
    lblSample.FontName = cmbFonts.Text
End Sub
Private Sub cmbFontSize_Click()
    If Val(cmbFontSize.Text) < 1 Or _
    Val(cmbFontSize.Text) > 1638 Then Exit Sub
    lblSample.FontSize = Val(cmbFontSize.Text)
End Sub
Private Sub Bold1_Click()
    If Bold1.Value = 1 Then
        lblSample.FontBold = True
    Else
        lblSample.FontBold = False
    End If
End Sub
Private Sub Italic2_Click()
    If Italic2.Value = 1 Then
        lblSample.FontItalic = True
    Else
        lblSample.FontItalic = False
    End If
End Sub
Private Sub Underline3_Click()
    If Underline3.Value = 1 Then
        lblSample.FontUnderline = True
    Else
        lblSample.FontUnderline = False
    End If
End Sub
Public Sub Change_Font(ctl As Control)
    With ctl
        .FontName = lblSample.FontName
        .FontSize = lblSample.FontSize
        .FontBold = lblSample.FontBold
        .FontItalic = lblSample.FontItalic
        .FontUnderline = lblSample.FontUnderline
        .ForeColor = lblSample.ForeColor
        
    End With
End Sub
    
    
    
