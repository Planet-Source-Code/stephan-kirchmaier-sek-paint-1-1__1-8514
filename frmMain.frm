VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "SEK Paint 1.0"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   8175
   LinkTopic       =   "Form1"
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   545
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picBCol 
      Height          =   255
      Left            =   7680
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   10
      Top             =   6600
      Width           =   375
   End
   Begin VB.PictureBox picFCol 
      Height          =   255
      Left            =   7200
      ScaleHeight     =   195
      ScaleWidth      =   315
      TabIndex        =   9
      Top             =   6600
      Width           =   375
   End
   Begin ComctlLib.ProgressBar pg1 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   6120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame fratemp 
      Height          =   15
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   8175
   End
   Begin VB.HScrollBar scrHorz 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   5760
      Width           =   7695
   End
   Begin VB.VScrollBar scrVert 
      Height          =   5655
      Left            =   7800
      TabIndex        =   3
      Top             =   120
      Width           =   255
   End
   Begin VB.PictureBox picBack 
      Height          =   5655
      Left            =   120
      ScaleHeight     =   373
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   0
      Top             =   120
      Width           =   7695
      Begin VB.PictureBox picZoom 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   375
         Left            =   5040
         ScaleHeight     =   21
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   15
         Top             =   3840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picFlip 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   5040
         ScaleHeight     =   435
         ScaleWidth      =   2115
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picPrint 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   5040
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   13
         Top             =   2640
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picUndoTools 
         AutoRedraw      =   -1  'True
         Height          =   495
         Left            =   5040
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picNew 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   5040
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   11
         Top             =   1320
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picUndo 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   5040
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   7
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSComDlg.CommonDialog comDiag 
         Left            =   4320
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox picTemp 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   495
         Left            =   5040
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   141
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.PictureBox picMain 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   3900
         Left            =   0
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   256
         TabIndex        =   1
         Top             =   0
         Width           =   3900
      End
   End
   Begin VB.Label lblCoords 
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   6480
      Width           =   2175
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuPencil 
         Caption         =   "&Pencil"
      End
      Begin VB.Menu mnuStar 
         Caption         =   "&Star"
      End
      Begin VB.Menu mnuHorzLine 
         Caption         =   "&Horizontal Line"
      End
      Begin VB.Menu mnuVertLine 
         Caption         =   "&Vertical Line"
      End
      Begin VB.Menu mnuCrossND 
         Caption         =   "Cro&ss"
      End
      Begin VB.Menu mnuCross 
         Caption         =   "Diagonal C&ross"
      End
      Begin VB.Menu mnuUDefPolygon 
         Caption         =   "User-defined Polygon"
      End
      Begin VB.Menu mnuDiagLineRL 
         Caption         =   "Diag&onal Line (\)"
      End
      Begin VB.Menu mnuDiagLineLR 
         Caption         =   "Dia&gonal Line (/)"
      End
      Begin VB.Menu mnuText 
         Caption         =   "Te&xt"
      End
      Begin VB.Menu mnuStLine 
         Caption         =   "&Straight Line"
      End
      Begin VB.Menu mnuBrush 
         Caption         =   "&Brush"
      End
      Begin VB.Menu mnuErase 
         Caption         =   "E&rase"
      End
      Begin VB.Menu mnuFRect 
         Caption         =   "Fi&lled Rect"
      End
      Begin VB.Menu mnuFCircle 
         Caption         =   "&Filled Circle"
      End
      Begin VB.Menu mnuCircle 
         Caption         =   "&Circle"
      End
      Begin VB.Menu mnuRect 
         Caption         =   "&Rect"
      End
      Begin VB.Menu mnuPolygon 
         Caption         =   "&Polygon"
      End
      Begin VB.Menu mnuFillReg 
         Caption         =   "&Fill Regions"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoTools 
         Caption         =   "Und&o"
      End
   End
   Begin VB.Menu mnuProps 
      Caption         =   "&Properties"
      Begin VB.Menu mnuSetDW 
         Caption         =   "&Set the DrawWidth"
      End
      Begin VB.Menu mnuDS 
         Caption         =   "&DrawStyle"
         Begin VB.Menu mnuDFilled 
            Caption         =   "&Filled"
         End
         Begin VB.Menu mnuDLine 
            Caption         =   "&Line"
         End
         Begin VB.Menu mnuDPoint 
            Caption         =   "&Point"
         End
         Begin VB.Menu mnuDLinPoint 
            Caption         =   "Li&ne-Point"
         End
         Begin VB.Menu mnuDLinPointPt 
            Caption         =   "Line&-Point-Point"
         End
      End
      Begin VB.Menu mnuFStyle 
         Caption         =   "&FillStyle"
         Begin VB.Menu mnuFilled 
            Caption         =   "&Filled"
         End
         Begin VB.Menu mnuFHorzLine 
            Caption         =   "&Horizontal Line"
         End
         Begin VB.Menu mnuFVertLine 
            Caption         =   "&Vertical Line"
         End
         Begin VB.Menu mnuFDiagonalRL 
            Caption         =   "Diagonal (&\)"
         End
         Begin VB.Menu mnuFDiagonalLR 
            Caption         =   "Diagonal (&/)"
         End
         Begin VB.Menu mnuFCross 
            Caption         =   "&Cross"
         End
         Begin VB.Menu mnuFDiagCross 
            Caption         =   "Diagonal Cr&oss"
         End
      End
   End
   Begin VB.Menu mnuFilters 
      Caption         =   "&Filters"
      Begin VB.Menu mnuEmboss 
         Caption         =   "&Emboss"
      End
      Begin VB.Menu mnuSharpen 
         Caption         =   "&Sharpen"
      End
      Begin VB.Menu mnuDiffuse 
         Caption         =   "&Diffuse"
      End
      Begin VB.Menu mnuRects 
         Caption         =   "&Rects"
      End
      Begin VB.Menu mnuBright 
         Caption         =   "&Brightness"
      End
      Begin VB.Menu mnuIce 
         Caption         =   "&Ice"
      End
      Begin VB.Menu mnuDark 
         Caption         =   "&Dark"
      End
      Begin VB.Menu mnuHeat 
         Caption         =   "&Heat"
      End
      Begin VB.Menu mnuStrange 
         Caption         =   "&Strange"
      End
      Begin VB.Menu mnuAqua 
         Caption         =   "&Aqua"
      End
      Begin VB.Menu mnuNight 
         Caption         =   "&Night"
      End
      Begin VB.Menu mnuAfrika 
         Caption         =   "&Afrika"
      End
      Begin VB.Menu mnuBlur 
         Caption         =   "&Blur"
      End
      Begin VB.Menu mnuInvert 
         Caption         =   "&Invert"
      End
      Begin VB.Menu mnuBAndW 
         Caption         =   "&Greyscale"
      End
      Begin VB.Menu mnuComic 
         Caption         =   "C&omic"
      End
      Begin VB.Menu mnuBaW 
         Caption         =   "B&lack and White"
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndo 
         Caption         =   "&Undo Filter"
      End
   End
   Begin VB.Menu mnuEffects 
      Caption         =   "&Effects"
      Begin VB.Menu mnuFlip1 
         Caption         =   "Flip&1"
      End
      Begin VB.Menu mnuFlip2 
         Caption         =   "Flip&2"
      End
      Begin VB.Menu mnuFlip3 
         Caption         =   "Flip&3"
      End
      Begin VB.Menu mnuRepCol 
         Caption         =   "&Replace Color"
      End
      Begin VB.Menu mnuWave 
         Caption         =   "&Wave"
      End
      Begin VB.Menu mnuHammer 
         Caption         =   "&Hammer"
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "&Zoom"
         Begin VB.Menu mnuZoomIn 
            Caption         =   "&Zoom In"
            Begin VB.Menu mnuP25 
               Caption         =   "+25 %"
            End
            Begin VB.Menu mnuP50 
               Caption         =   "+50%"
            End
            Begin VB.Menu mnuP75 
               Caption         =   "+75%"
            End
         End
         Begin VB.Menu mnuZoomOut 
            Caption         =   "Z&oom Out"
            Begin VB.Menu mnuM25 
               Caption         =   "-25%"
            End
            Begin VB.Menu mnuM50 
               Caption         =   "-50%"
            End
            Begin VB.Menu mnuM75 
               Caption         =   "-75%"
            End
         End
      End
      Begin VB.Menu mnuSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUndoEffects 
         Caption         =   "&Undo"
      End
   End
   Begin VB.Menu mnuInfo 
      Caption         =   "&Info"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    cX = picMain.ScaleWidth
    cY = picMain.ScaleHeight
    scrHorz.Value = 0
    scrHorz.Max = picMain.Width - 5
    scrVert.Value = 0
    scrVert.Max = picMain.Height - 5
    tColors.lBCol = vbBlue
    tColors.lFCol = vbBlack
    picFCol.BackColor = vbBlack
    picBCol.BackColor = vbBlue
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblCoords.Caption = vbNullString
End Sub

Private Sub Form_Resize()
    
    On Error GoTo Fa
    
    With picBack
        .Top = 8
        .Left = 8
        .Width = frmMain.ScaleWidth - 28
        .Height = frmMain.ScaleHeight - 60
    End With
    With scrHorz
        .Left = 8
        .Top = 10 + picBack.ScaleHeight
        .Width = picBack.ScaleWidth + 3
    End With
    With scrVert
        .Left = 8 + picBack.ScaleWidth
        .Top = 8
        .Height = picBack.ScaleHeight + 3
    End With
    With pg1
        .Top = scrHorz.Top + 20
        .Width = picBack.ScaleWidth + 17
    End With
    With lblCoords
        .Left = 8
        .Top = pg1.Top + pg1.Height + 1
    End With
    With picBCol
        .Left = frmMain.ScaleWidth - 65
        .Top = frmMain.ScaleHeight - 16
    End With
    With picFCol
        .Left = frmMain.ScaleWidth - 33
        .Top = frmMain.ScaleHeight - 16
    End With
    fratemp.Width = frmMain.ScaleWidth
Fa:

End Sub

Private Sub mnuAbout_Click()
    MsgBox "Made by Stephan Kirchmaier in Y2K" & vbCr & "Please vote for me ;-)", vbInformation, "About"
End Sub

Private Sub mnuBaW_Click()
    Dim col As Long
    
    Call Save
    For i = 0 To picMain.Width
        For j = 0 To picMain.Height
            col = GetPixel(picMain.hdc, i, j)
            r = col Mod 256
            g = (col Mod 256) \ 256
            b = col \ 256 \ 256
            
            If r < 200 And g < 200 And b < 200 Then
                col = vbBlack
            Else
                col = vbWhite
            End If
            SetPixel picMain.hdc, i, j, col
        Next j
        pg1.Value = i * 100 \ (picMain.Width - 1)
    Next i
    pg1.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuBrush_Click()
    curTools = Tools.sBrush
End Sub

Private Sub mnuCircle_Click()
    curTools = Tools.sCircle
End Sub

Private Sub mnuComic_Click()
    Dim col As Long
    
    Call Save
    For i = 0 To picMain.Width
        For j = 0 To picMain.Height
            col = GetPixel(picMain.hdc, i, j)
            r = Abs(col Mod 256)
            g = Abs((col \ 256) Mod 256)
            b = Abs(col \ 256 \ 256)
            r = Abs(r * (g - b + g + r)) / 256
            g = Abs(r * (b - g + b + r)) / 256
            b = Abs(g * (b - g + b + r)) / 256
            col = RGB(r, g, b)
            r = Abs(col Mod 256)
            g = Abs((col \ 256) Mod 256)
            b = Abs(col \ 256 \ 256)
            r = (r + g + b) / 3
            col = RGB(r, r, r)
            SetPixel picMain.hdc, i, j, col
        Next j
        pg1.Value = i * 100 \ (picMain.Width - 1)
    Next i
    pg1.Value = 0
    picMain.Refresh
    
End Sub

Private Sub mnuCross_Click()
    curTools = Tools.sCross
End Sub

Private Sub mnuCrossND_Click()
    curTools = Tools.sCrossND
End Sub

Private Sub mnuDFilled_Click()
    picMain.DrawStyle = 0
End Sub

Private Sub mnuDiagLineLR_Click()
    curTools = Tools.sDiagLineLR
End Sub

Private Sub mnuDiagLineRL_Click()
    curTools = Tools.sDiagLineRL
End Sub

Private Sub mnuDLine_Click()
    picMain.DrawStyle = 1
End Sub

Private Sub mnuDLinPoint_Click()
    picMain.DrawStyle = 3
End Sub

Private Sub mnuDLinPointPt_Click()
    picMain.DrawStyle = 4
End Sub

Private Sub mnuDPoint_Click()
    picMain.DrawStyle = 2
End Sub

Private Sub mnuErase_Click()
    curTools = Tools.sErase
End Sub

Private Sub mnuFCircle_Click()
    curTools = Tools.sFCircle
End Sub

Private Sub mnuFCross_Click()
    propFillStyle = 6
End Sub

Private Sub mnuFDiagCross_Click()
    propFillStyle = 7
End Sub

Private Sub mnuFDiagonalLR_Click()
    propFillStyle = 5
End Sub

Private Sub mnuFDiagonalRL_Click()
    propFillStyle = 4
End Sub

Private Sub mnuFHorzLine_Click()
    propFillStyle = 2
End Sub

Private Sub mnuFilled_Click()
    propFillStyle = 0
End Sub

Private Sub mnuFillReg_Click()
    curTools = Tools.sFillRegions
End Sub

Private Sub mnuFlip1_Click()
    Call Save
    Call PrepFlip
    
    picMain.PaintPicture picFlip.Picture, 0, picMain.ScaleHeight - 1, picMain.ScaleWidth, -picMain.ScaleHeight, , , , , vbSrcCopy

End Sub

Private Sub mnuFlip2_Click()
    Call Save
    
    Call PrepFlip
    picMain.PaintPicture picFlip.Picture, picMain.ScaleWidth - 1, 0, -picMain.ScaleWidth, picMain.ScaleHeight, , , , , vbSrcCopy

End Sub

Private Sub mnuFlip3_Click()
    Call Save
    Call PrepFlip
    picMain.PaintPicture picFlip.Picture, picMain.ScaleWidth - 1, picMain.ScaleHeight - 1, -picMain.ScaleWidth, -picMain.ScaleHeight, , , , , vbSrcCopy
    
End Sub

Private Sub mnuFRect_Click()
    curTools = Tools.sFRect
End Sub

Private Sub mnuFVertLine_Click()
    propFillStyle = 3
End Sub

Private Sub mnuHammer_Click()
    curTools = Tools.sHammer
End Sub

Private Sub mnuHeat_Click()
    Dim bNo As Boolean
    Dim TColW As Long
    
    Call Save
    For i = 0 To cX
        For j = 0 To cY
            TColW = GetPixel(picMain.hdc, i, j)
            r = TColW Mod 256
            g = (TColW \ 256) Mod 256
            b = TColW \ 256 \ 256
            
            r = Abs(((r ^ 2) / ((b + g) + 10)) * 128)
            b = Abs(((b ^ 2) / ((g + r) + 10)) * 128)
            g = Abs(((g ^ 2) / ((r + b) + 10)) * 128)
nOK:
                If r > 32767 Then
                    r = r - 32767
                ElseIf g > 32767 Then
                    g = g - 32767
                ElseIf b > 32767 Then
                    b = b - 32767
                End If
                If r > 32767 Or g > 32767 Or b > 32767 Then
                    GoTo nOK
                End If
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    pg1.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuHorzLine_Click()
    curTools = Tools.sHorzLine
End Sub

Private Sub mnuIce_Click()
    Dim TColI As Long
    
    Call Save
    For i = 0 To cX
        For j = 0 To cY
            TColI = GetPixel(picMain.hdc, i, j)
            r = TColI Mod 256
            g = (TColI \ 256) Mod 256
            b = TColI \ 256 \ 256
            r = Abs((r - g - b) * 1.5)
            g = Abs((g - b - r) * 1.5)
            b = Abs((b - r - g) * 1.5)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    pg1.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuM25_Click()
    Call Zoom(0.25)
End Sub

Private Sub mnuM50_Click()
    Call Zoom(0.5)
End Sub

Private Sub mnuM75_Click()
    Call Zoom(0.75)
End Sub

Private Sub mnuNew_Click()
    picMain.Refresh
    picNew.Height = picMain.Height
    picNew.Width = picMain.Width
    picMain.Picture = picNew.Image
End Sub

Private Sub mnuOpen_Click()
    Dim sFName As String
    
    On Error Resume Next
    comDiag.Filter = "*.bmp;*.jpg;*.gif;*.wmf;"
    comDiag.ShowOpen
    sFName = comDiag.filename
    
    If sFName = "" Then Exit Sub
    
    If Not FileExist(sFName) Then
        MsgBox "File doesn't exist.", vbCritical, "Error"
        Exit Sub
    End If
    
    picMain.Picture = LoadPicture(sFName)
    If Err Then
        MsgBox "This is not a valid picture!", vbCritical, "Error"
        Exit Sub
    End If
    Call ResizePicBoxes
    cX = picMain.ScaleWidth
    cY = picMain.ScaleHeight
End Sub

Private Sub mnuP25_Click()
    Call Zoom(1.25)
End Sub

Private Sub mnuP50_Click()
    Call Zoom(1.5)
End Sub

Private Sub mnuP75_Click()
    Call Zoom(1.75)
End Sub

Private Sub mnuPencil_Click()
    curTools = Tools.sPencil
End Sub

Private Sub mnuPolygon_Click()
    curTools = Tools.sPolygon
End Sub

Private Sub mnuPrint_Click()
    frmPrint.Show
End Sub

Private Sub mnuRect_Click()
    curTools = Tools.sRect
End Sub

Private Sub mnuRects_Click()
    Dim tColR1 As Long, tColR2 As Long, tColR3 As Long, tColR4 As Long, tColR5 As Long
    
    Call Save
    For i = 0 To cX
        For j = 0 To cY
            tColR1 = GetPixel(picMain.hdc, i, j)
            tColR2 = GetPixel(picMain.hdc, i + 1, j)
            tColR3 = GetPixel(picMain.hdc, i - 1, j)
            tColR4 = GetPixel(picMain.hdc, i, j + 1)
            tColR5 = GetPixel(picMain.hdc, i, j - 1)
            SetPixel picMain.hdc, i, j, (Abs(tColR1) - (Abs(tColR2 + tColR3 + tColR4 + tColR5) / 4))
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    pg1.Value = 0
    picMain.Refresh
End Sub

Private Sub mnuRepCol_Click()
    Call Save
    MsgBox "At first, you must choose the color you want to replace on the picture. Then you must choose the ""replacing-color"".", vbInformation, "SEK - Paint 1.0"
    FirstChoose = True
    curTools = Tools.sReplaceColor
End Sub

Private Sub mnuSave_Click()
    Dim sFName As String
    
    On Error Resume Next
    comDiag.CancelError = True
    comDiag.Filter = "*.bmp"
    comDiag.ShowSave
    If Err Then Exit Sub
    sFName = comDiag.filename
    
    If FileExist(sFName) Then
        MsgBox "The File already exist!", vbCritical, "Error"
        Exit Sub
    End If
    
    If (LCase$(Right$(sFName, 4)) = ".bmp") Then
        SavePicture picMain.Image, sFName
    Else
        sFName = sFName & ".bmp"
        SavePicture picMain.Image, sFName
    End If
    
End Sub

Private Sub mnuSetDW_Click()
    Dim iCool As Integer
    
    iCool = InputBox("Type in the desired Draw Width!", "SEK - Paint 1.0", picMain.DrawWidth)
    If Not IsNumeric(iCool) Then
        MsgBox "You must type in a valid number!", vbCritical, "Error"
        Exit Sub
    End If
    picMain.DrawWidth = iCool
End Sub

Private Sub mnuStar_Click()
    curTools = Tools.sStar
End Sub

Private Sub mnuStLine_Click()
    curTools = Tools.sStLine
End Sub

Private Sub mnuText_Click()
    curTools = Tools.sText
    frmText.Show
End Sub

Private Sub mnuUDefPolygon_Click()
    curTools = Tools.sUDefPolygon
End Sub

Private Sub mnuUndo_Click()
    Call UndoFilters
End Sub

Private Sub mnuUndoEffects_Click()
    Call UndoFilters
End Sub

Private Sub mnuUndoTools_Click()
    On Error Resume Next
    
    Set picMain.Picture = picUndoTools.Image
End Sub

Private Sub mnuVertLine_Click()
    curTools = Tools.sVertLine
End Sub

Private Sub mnuWave_Click()
    Dim i As Long, j As Long
    Dim sw As Long, sh As Long
    Dim coli() As Long, posy() As Double
    
    Call Save
    sw = picMain.Width
    sh = picMain.Height
    
    ReDim coli(sw, sh)
    ReDim posy(sw, sh)
    
    For i = 0 To sw
        For j = 0 To sh
            coli(i, j) = GetPixel(picMain.hdc, i, j)
            posy(i, j) = Sin(i) * 6 + (j - 3)
        Next j
        pg1.Value = i * 100 \ (sw - 1)
    Next i
    For i = 0 To sw
        For j = 0 To sh
            picMain.PSet (i, posy(i, j)), coli(i, j)
        Next j
        pg1.Value = i * 100 \ (sw - 1)
    Next i
    pg1.Value = 0
    
End Sub

Private Sub picBack_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        picMain.Width = x
        picMain.Height = y
        Call ResizePicBoxes
    End If
End Sub

Private Sub picBack_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblCoords.Caption = vbNullString
End Sub

Private Sub picBCol_Click()
    On Error Resume Next
    
    comDiag.CancelError = True
    comDiag.ShowColor
    If Not Err Then
        picBCol.BackColor = comDiag.Color
        tColors.lBCol = comDiag.Color
    End If
End Sub

Private Sub picFCol_Click()
    On Error Resume Next
    
    comDiag.CancelError = True
    comDiag.ShowColor
    If Not Err Then
        picFCol.BackColor = comDiag.Color
        tColors.lFCol = comDiag.Color
    End If
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sTmp As String
    Dim points(2) As cPoint, cCount As Double, cDetails As Double
    Dim col As Long, r As Long, a As Double, u As Long, v As Long
    
    sX = x
    sY = y
    If Button = 1 Then
        Call savPic
        If curTools = Tools.sCircle Or curTools = Tools.sStLine Or curTools = Tools.sFCircle Or curTools = Tools.sFRect Or curTools = Tools.sRect Then
            picMain.PSet (sX, sY), vbBlack
            Exit Sub
        End If
        
        If curTools = Tools.sPolygon Then
            If Not stat Then
                stat = True
                curX = x
                curY = y
            End If
            sX = x
            sY = y
            If Shift = 1 Then
                picMain.Line (wX1, wY1)-(curX, curY), tColors.lFCol
                stat1 = False
                stat = False
                Exit Sub
            End If
            If Not stat1 Then
                picMain.PSet (x, y), tColors.lFCol
                wX1 = x
                wY1 = y
                stat1 = True
            Else
                picMain.Line (wX1, wY1)-(x, y), tColors.lFCol
            End If
            Exit Sub
        End If
        
        If curTools = Tools.sUDefPolygon Then
            upX = x
            upY = y
        End If
        
        If curTools = Tools.sFillRegions Then
            Call Filling(picMain.Point(x, y), propFillStyle, x, y)
            
        End If
        
        If curTools = Tools.sText Then
            picMain.CurrentX = x
            picMain.CurrentY = y
            sTmp = InputBox("Type in the Text!", "SEK - Paint 1.0")
            picMain.Print sTmp
        End If
        
        If curTools = Tools.sHammer Then
            On Error Resume Next
            
            Call Save
            r = InputBox("Please type in the radius!  (1-150)", "Bézier")
            If Err Or r < 0 Or r > 150 Then
                MsgBox "Please type in a number between 1 and 150!", vbCritical, "Error"
                Exit Sub
            End If
            a = ((r / 50) * 360)
        
            For i = 0 To a
                points(0).cX = r * Cos(i) + x
                points(0).cY = r * Sin(i) + y
                points(1).cX = x
                points(1).cY = y
                points(2).cX = x
                points(2).cY = y
        
                cCount = 0
                cDetails = 1 / (r * 2)
        
                Do
                    col = GetPixel(picMain.hdc, points(0).cX, points(0).cY)
                    u = points(0).cX * (cCount * cCount) + points(1).cX * (2 * cCount * (1 - cCount)) + points(2).cX * ((1 - cCount) * (1 - cCount))
                    v = points(0).cY * (cCount * cCount) + points(1).cY * (2 * cCount * (1 - cCount)) + points(2).cY * ((1 - cCount) * (1 - cCount))
                    picMain.PSet (sX, sY), col
                    SetPixel picMain.hdc, u, v, col
                    cCount = cCount + cDetails
                Loop While cCount <= 1
                pg1.Value = i * 100 \ (a - 1)
            Next i
            pg1.Value = 0
        End If
        
        If curTools = Tools.sReplaceColor Then
            If FirstChoose Then
                tmpCol1 = picMain.Point(x, y)
                FirstChoose = False
            Else
                tmpCol2 = picMain.Point(x, y)
                Call ReplaceColor
            End If
        End If
    End If
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim tmpbX As Integer, tmpbY As Integer
    
    lblCoords.Caption = x & " X " & y
    If Button = 1 Then
        Select Case curTools
            Case Tools.sPencil:
                picMain.Line (sX, sY)-(x, y), tColors.lFCol
                sX = x
                sY = y
                
            Case Tools.sStar:
                picMain.Line (sX, sY)-(x, y), tColors.lFCol
            
            Case Tools.sErase:
                picMain.Line (sX, sY)-(x, y), vbWhite
                sX = x
                sY = y
            
            Case Tools.sCross:
                picMain.Line (x - 5, y - 5)-(x + 5, y + 5), tColors.lFCol
                picMain.Line (x + 5, y - 5)-(x - 5, y + 5), tColors.lFCol
            
            Case Tools.sCrossND:
                picMain.Line (x - 5, y)-(x + 5, y), tColors.lFCol
                picMain.Line (x, y - 5)-(x, y + 5), tColors.lFCol
                
            Case Tools.sDiagLineLR:
                picMain.Line (x, y)-(x - 5, y + 5), tColors.lFCol
                
            Case Tools.sDiagLineRL:
                picMain.Line (x, y)-(x + 5, y + 5), tColors.lFCol
            
            Case Tools.sUDefPolygon:
                picMain.Line (sX, sY)-(x, y), tColors.lFCol
                sX = x
                sY = y
            
            Case Tools.sHorzLine:
                picMain.Line (x - 5, y)-(x + 5, y), tColors.lFCol
                
            Case Tools.sVertLine:
                picMain.Line (x, y - 5)-(x, y + 5), tColors.lFCol
        
            Case Tools.sBrush:
                For i = 0 To 25
                    tmpbX = Int(Rnd * 14 - 7)
                    tmpbY = Int(Rnd * 14 - 7)
                    picMain.PSet (x + tmpbX, y + tmpbY), tColors.lFCol
                Next i
            
        End Select
    End If
    wX = x
    wY = y
End Sub

Private Sub picMain_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim r
    
    If Button = 1 Then
        If curTools = Tools.sCircle Then
            r = Sqr((sX - x) * (sX - x) + (sY - y) * (sY - y))
            picMain.Circle (sX, sY), r, tColors.lFCol
            picMain.PSet (sX, sY), picMain.Point(sX + 1, sY)
            
        ElseIf curTools = Tools.sRect Then
            picMain.Line (sX, sY)-(x, y), tColors.lFCol, B
            
        ElseIf curTools = Tools.sStLine Then
            picMain.Line (sX, sY)-(x, y), tColors.lFCol
        
        ElseIf curTools = Tools.sFCircle Then
            picMain.FillStyle = 0
            picMain.FillColor = tColors.lBCol
            r = Sqr((sX - x) * (sX - x) + (sY - y) * (sY - y))
            picMain.Circle (sX, sY), r, tColors.lFCol
            picMain.FillStyle = 1
        
        ElseIf curTools = Tools.sFRect Then
            picMain.Line (sX, sY)-(x, y), tColors.lBCol, BF
            
        ElseIf curTools = Tools.sPolygon Then
            picMain.Line (sX, sY)-(wX, wY), tColors.lFCol
            wX1 = x
            wY1 = y
            
        ElseIf curTools = Tools.sUDefPolygon Then
            picMain.Line (upX, upY)-(x, y), tColors.lFCol

        End If
        
    End If
End Sub

Private Sub scrHorz_Change()
    picMain.Left = -scrHorz.Value
End Sub

Private Sub scrVert_Change()
    picMain.Top = -scrVert.Value
End Sub
Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuInvert_Click()
    Call ResizePicBoxes
    Call Save
    SavePicture picMain.Image, App.Path & "\Invert.tmp"
    
    With picTemp
        picMain.Picture = LoadPicture(App.Path & "\Invert.tmp")
        .Picture = picMain.Picture
        picMain.PaintPicture .Picture, 0, 0, .Width, .Height, 0, 0, .Width, .Height, vbNotSrcCopy
    End With
    Kill App.Path & "\Invert.tmp"
End Sub


Private Sub mnuBAndW_Click()
    Dim c As Integer
    
    Call Save
    Call PrepareImg
    For i = 0 To cX
        For j = 0 To cY
            c = larrCol(0, i, j) * 0.3 + larrCol(1, i, j) * 0.59 + larrCol(2, i, j) * 0.11
            SetPixel picMain.hdc, i, j, RGB(c, c, c)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuEmboss_Click()
   
    Call Save
    Call PrepareImg
    For i = 0 To cX - 1
        For j = 0 To cY - 1
            r = Abs(larrCol(0, i, j) - larrCol(0, i + 1, j + 1) + 128)
            g = Abs(larrCol(1, i, j) - larrCol(1, i + 1, j + 1) + 128)
            b = Abs(larrCol(2, i, j) - larrCol(2, i + 1, j + 1) + 128)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuSharpen_Click()
    
    Call Save
    Call PrepareImg
    For i = 1 To cX
        For j = 1 To cY
            r = larrCol(0, i, j) + 0.5 * (larrCol(0, i, j) - larrCol(0, i - 1, j - 1))
            g = larrCol(1, i, j) + 0.5 * (larrCol(1, i, j) - larrCol(1, i - 1, j - 1))
            b = larrCol(2, i, j) + 0.5 * (larrCol(2, i, j) - larrCol(2, i - 1, j - 1))
            
            If r > 255 Then r = 255
            If r < 0 Then r = 0
            If g > 255 Then g = 255
            If g < 0 Then g = 0
            If b > 255 Then b = 255
            If b < 0 Then b = 0

            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuBright_Click()
    Dim c As Long

    Call Save
    Call PrepareImg
    For i = 0 To cX
        For j = 0 To cY
            c = Abs((larrCol(0, i, j) + larrCol(1, i, j) + larrCol(2, i, j)) \ 3)
            r = Abs(larrCol(0, i, j) + c)
            g = Abs(larrCol(1, i, j) + c)
            b = Abs(larrCol(2, i, j) + c)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuDiffuse_Click()
    Dim nP1 As Integer, nP2 As Integer, nP3 As Integer
    
    Call Save
    Call PrepareImg
    For i = 2 To cX - 3
        For j = 2 To cY - 3
            nP1 = Int(Rnd * 5 - 2)
            nP2 = Int(Rnd * 5 - 2)
            nP3 = Int(Rnd * 5 - 2)
            r = Abs(larrCol(0, i, j + nP1))
            g = Abs(larrCol(1, i + nP2, j))
            b = Abs(larrCol(2, i + nP3, j + nP3))
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuDark_Click()
    
    Call Save
    Call PrepareImg
    For i = 0 To cX
        For j = 0 To cY
            r = Abs(larrCol(0, i, j) - 64)
            g = Abs(larrCol(1, i, j) - 64)
            b = Abs(larrCol(2, i, j) - 64)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuStrange_Click()
    
    Call Save
    Call PrepareImg
    For i = 0 To cX
        For j = 0 To cY
            If (larrCol(1, i, j) = 0) Or (larrCol(2, i, j) = 0) Then
                larrCol(1, i, j) = 1
                larrCol(2, i, j) = 1
            End If
            r = Abs(Sin(Atn(larrCol(1, i, j) / larrCol(2, i, j))) * 125 + 20)
            g = Abs(Sin(Atn(larrCol(0, i, j) / larrCol(2, i, j))) * 125 + 20)
            b = Abs(Sin(Atn(larrCol(0, i, j) / larrCol(1, i, j))) * 125 + 20)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuAqua_Click()
    Dim tColQ As Long
    
    Call Save
    For i = 0 To cX
        For j = 0 To cY
            tColQ = GetPixel(picMain.hdc, i, j)
            r = tColQ Mod 256
            g = (tColQ \ 256) Mod 256
            b = tColQ \ 256 \ 256
            r = (g - b) ^ 2 / 125
            g = (r - b) ^ 2 / 125
            b = (r - g) ^ 2 / 125
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuNight_Click()
    
    Call Save
    Call PrepareImg
    For i = 0 To cX
        For j = 0 To cY
            r = Abs((larrCol(0, i, j) * larrCol(0, i, j)) / 256)
            g = Abs((larrCol(1, i, j) * larrCol(1, i, j)) / 256)
            b = Abs((larrCol(2, i, j) * larrCol(2, i, j)) / 256)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuAfrika_Click()
    Dim TColA
    
    Call Save
    For i = 0 To cX
        For j = 0 To cY
            TColA = GetPixel(picMain.hdc, i, j)
            r = TColA Mod 256
            g = (TColA \ 256) Mod 256
            b = TColA \ 256 \ 256
            r = Abs((g * b) / 256)
            g = Abs((b * r) / 256)
            b = Abs((r * g) / 256)
            SetPixel picMain.hdc, i, j, RGB(r, g, b)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub

Private Sub mnuBlur_Click()
    
    Call Save
    Call PrepareImg
    For i = 1 To cX - 1
        For j = 1 To cY - 1
            r = Abs(larrCol(0, i - 1, j - 1) + larrCol(0, i, j - 1) + larrCol(0, i + 1, j - 1) + larrCol(0, i - 1, j) + larrCol(0, i, j) + larrCol(0, i + 1, j) + larrCol(0, i - 1, j + 1) + larrCol(0, i, j + 1) + larrCol(0, i + 1, j + 1))
            g = Abs(larrCol(1, i - 1, j - 1) + larrCol(1, i, j - 1) + larrCol(1, i + 1, j - 1) + larrCol(1, i - 1, j) + larrCol(1, i, j) + larrCol(1, i + 1, j) + larrCol(1, i - 1, j + 1) + larrCol(1, i, j + 1) + larrCol(1, i + 1, j + 1))
            b = Abs(larrCol(2, i - 1, j - 1) + larrCol(2, i, j - 1) + larrCol(2, i + 1, j - 1) + larrCol(2, i - 1, j) + larrCol(2, i, j) + larrCol(2, i + 1, j) + larrCol(2, i - 1, j + 1) + larrCol(2, i, j + 1) + larrCol(2, i + 1, j + 1))
            SetPixel picMain.hdc, i, j, RGB(r / 10, g / 10, b / 10)
        Next j
        pg1.Value = i * 100 \ (cX - 1)
    Next i
    picMain.Refresh
    pg1.Value = 0
End Sub
