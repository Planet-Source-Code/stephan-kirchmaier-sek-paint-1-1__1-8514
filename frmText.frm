VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Set the Text"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CheckBox chkStrikeTh 
      Caption         =   "Strikethrough"
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   3360
      Width           =   1335
   End
   Begin VB.CheckBox chkUnderlined 
      Caption         =   "Underlined Text"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   3360
      Width           =   1455
   End
   Begin VB.OptionButton optBoldItalic 
      Caption         =   "Bold-Italic"
      Height          =   255
      Left            =   3720
      TabIndex        =   9
      Top             =   3000
      Width           =   1095
   End
   Begin VB.OptionButton optItalic 
      Caption         =   "Italic"
      Height          =   255
      Left            =   2640
      TabIndex        =   8
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton optBold 
      Caption         =   "Bold"
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   3000
      Width           =   735
   End
   Begin VB.OptionButton optStandard 
      Caption         =   "Standard"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   2280
      Width           =   1695
   End
   Begin VB.ComboBox cboSize 
      Height          =   315
      ItemData        =   "frmText.frx":0000
      Left            =   3000
      List            =   "frmText.frx":0002
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   360
      Width           =   1695
   End
   Begin VB.ListBox lstFonts 
      Height          =   2400
      ItemData        =   "frmText.frx":0004
      Left            =   240
      List            =   "frmText.frx":0006
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label lblFontSize 
      AutoSize        =   -1  'True
      Caption         =   "Font Size:"
      Height          =   195
      Left            =   3000
      TabIndex        =   3
      Top             =   120
      Width           =   705
   End
   Begin VB.Label lblFontsChoose 
      AutoSize        =   -1  'True
      Caption         =   "Choose your Font:"
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1290
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmText.Hide
End Sub

Private Sub cmdOK_Click()
    With frmMain.picMain
        .FontName = lstFonts.List(lstFonts.ListIndex)
        .FontSize = cboSize.List(cboSize.ListIndex)
        .FontBold = optBold
        .FontItalic = optItalic
        If optBoldItalic Then
            .FontBold = True
            .FontItalic = True
        End If
        .FontStrikethru = chkStrikeTh.Value
        .FontUnderline = chkUnderlined.Value
    End With
    frmText.Hide
    
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    lstFonts.Clear
    cboSize.Clear
    For i = 0 To Printer.FontCount - 1
        lstFonts.AddItem Printer.Fonts(i)
    Next i
    For i = 2 To 80 Step 2
        cboSize.AddItem CStr(i)
    Next i
    optStandard.Value = True
    cboSize.ListIndex = 0
    
End Sub
