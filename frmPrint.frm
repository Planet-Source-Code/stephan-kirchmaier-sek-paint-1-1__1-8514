VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmPrint 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Printing"
   ClientHeight    =   3300
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.TextBox txtLeft 
      Height          =   285
      Left            =   4800
      TabIndex        =   18
      Text            =   "10"
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox txtTop 
      Alignment       =   1  'Rechts
      Height          =   285
      Left            =   4800
      TabIndex        =   14
      Text            =   "10"
      Top             =   600
      Width           =   735
   End
   Begin VB.Frame fraPQu 
      Caption         =   "Print Quality"
      Height          =   2295
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   2415
      Begin VB.TextBox txtDpi 
         Alignment       =   1  'Rechts
         Height          =   285
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   10
         Top             =   1800
         Width           =   495
      End
      Begin VB.OptionButton optUDef 
         Caption         =   "User defined:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1275
      End
      Begin VB.OptionButton optBest 
         Caption         =   "Best Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1440
         Width           =   1215
      End
      Begin VB.OptionButton optBetter 
         Caption         =   "Better Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optGood 
         Caption         =   "Good Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.OptionButton optBad 
         Caption         =   "Bad Quality"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblDpi 
         AutoSize        =   -1  'True
         Caption         =   "dpi"
         Height          =   195
         Left            =   2040
         TabIndex        =   11
         Top             =   1920
         Width           =   210
      End
   End
   Begin ComctlLib.Slider sldNumCop 
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   327682
      Min             =   1
      Max             =   20
      SelStart        =   1
      Value           =   1
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3840
      TabIndex        =   2
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print!"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label lblPix2 
      AutoSize        =   -1  'True
      Caption         =   "in Pixels:"
      Height          =   195
      Left            =   3120
      TabIndex        =   17
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label lblLeft 
      AutoSize        =   -1  'True
      Caption         =   "Distance from the Left of the page"
      Height          =   195
      Left            =   3120
      TabIndex        =   16
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label lblPix1 
      AutoSize        =   -1  'True
      Caption         =   "in Pixels:"
      Height          =   195
      Left            =   3120
      TabIndex        =   15
      Top             =   600
      Width           =   615
   End
   Begin VB.Label lblTop 
      AutoSize        =   -1  'True
      Caption         =   "Distance from the Top of the page"
      Height          =   195
      Left            =   3120
      TabIndex        =   13
      Top             =   240
      Width           =   2430
   End
   Begin VB.Label lblNumCop 
      AutoSize        =   -1  'True
      Caption         =   "01"
      Height          =   195
      Left            =   5400
      TabIndex        =   12
      Top             =   2160
      Width           =   180
   End
   Begin VB.Label lblNumCopEx 
      AutoSize        =   -1  'True
      Caption         =   "Number of copies:"
      Height          =   195
      Left            =   3120
      TabIndex        =   0
      Top             =   1920
      Width           =   1290
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    frmPrint.Hide
End Sub

Private Sub cmdPrint_Click()
    Dim iQuality As Integer

    If Not IsNumeric(txtLeft.Text) Then
        MsgBox "Type in a number, LEFT.", vbCritical, "Error"
        Exit Sub
    End If
    If Not IsNumeric(txtTop.Text) Then
        MsgBox "Type in a number, TOP.", vbCritical, "Error"
        Exit Sub
    End If
    If optUDef And Not IsNumeric(txtDpi.Text) Then
        MsgBox "Type in a number, DPI.", vbCritical, "Error"
        Exit Sub
    End If
    
    If optBad.Value Then iQuality = -1
    If optGood.Value Then iQuality = -2
    If optBetter.Value Then iQuality = -3
    If optBest.Value Then iQuality = -4
    If optUDef.Value Then iQuality = txtDpi.Text
    
    With Printer
        .Copies = sldNumCop.Value
        .PrintQuality = iQuality
        .ScaleMode = vbPixels
        .Height = frmMain.picMain.ScaleHeight
        .Width = frmMain.picMain.ScaleWidth
        .ScaleLeft = txtLeft.Text
        .ScaleTop = txtTop.Text
    End With
    
    With frmMain.picPrint
        .Height = frmMain.picMain.ScaleHeight
        .Width = frmMain.picMain.ScaleWidth
        .Picture = frmMain.picMain.Image
    End With
    
    Printer.PaintPicture frmMain.picPrint.Picture, 0, 0
    Printer.EndDoc
    cmdCancel.Caption = "Close"
End Sub

Private Sub Form_Load()
    sldNumCop.Value = 1
    optGood.Value = True
End Sub

Private Sub sldNumCop_Scroll()
    If sldNumCop.Value < 10 Then
        lblNumCop.Caption = "0" & CStr(sldNumCop.Value)
    Else
        lblNumCop.Caption = CStr(sldNumCop.Value)
    End If
End Sub
