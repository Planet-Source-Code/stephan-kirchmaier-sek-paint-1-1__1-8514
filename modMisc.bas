Attribute VB_Name = "modMisc"
Option Explicit

Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long

Global cX As Long, cY As Long
Global i As Long, j As Long
Global tmpCol As Long
Global r As Long, g As Long, b As Long
Global larrCol() As Long
Global curTools As Integer
Global tColors As Colors
Global sX As Long, sY As Long
Global upX As Long, upY As Long
Global stat As Boolean, stat1 As Boolean
Global wX As Long, wX1 As Long, wY As Long, wY1 As Long
Global curX As Long, curY As Long
Global propFillStyle As Long
Global tmpCol1 As Long, tmpCol2 As Long, FirstChoose As Boolean

Public Type cPoint
    cX As Double
    cY As Double
End Type

Public Type Colors
    lBCol As Long
    lFCol As Long
End Type

Public Enum Tools
    sPencil = 0
    sFCircle = 1
    sFRect = 2
    sStar = 3
    sCircle = 4
    sPolygon = 5
    sRect = 6
    sErase = 7
    sStLine = 8
    sBrush = 9
    sFillRegions = 10
    sText = 11
    sDiagLineRL = 12
    sDiagLineLR = 13
    sUDefPolygon = 14
    sCross = 15
    sVertLine = 16
    sHorzLine = 17
    sCrossND = 18
    sReplaceColor = 19
    sHammer = 20
End Enum


Public Sub PrepareImg()
    ReDim larrCol(2, cX, cY)
    For i = 0 To cX
        For j = 0 To cY
            tmpCol = GetPixel(frmMain.picMain.hdc, i, j)
            r = tmpCol Mod 256
            g = (tmpCol / 256) Mod 256
            b = tmpCol / 256 / 256
            larrCol(0, i, j) = r
            larrCol(1, i, j) = g
            larrCol(2, i, j) = b
        Next j
        frmMain.pg1.Value = i * 100 \ (cX - 1)
    Next i
    frmMain.pg1.Value = 0
End Sub

Public Sub Save()
    frmMain.picUndo.Picture = frmMain.picMain.Image
End Sub

Public Sub UndoFilters()
    On Error Resume Next
    
    frmMain.picMain.PaintPicture frmMain.picUndo.Picture, 0, 0
End Sub

Public Sub ResizePicBoxes()
    Dim lw As Long, lh As Long
    
    With frmMain
        .picMain.Refresh
        lh = .picMain.Height
        lw = .picMain.Width
        .picUndo.Height = lh
        .picUndo.Width = lw
        .picTemp.Height = lh
        .picTemp.Width = lw
        .picNew.Height = lh
        .picNew.Width = lw
        .picUndoTools.Height = lh
        .picUndoTools.Width = lw
        .picPrint.Height = lh
        .picPrint.Width = lw
        .picFlip.Width = lw
        .picFlip.Height = lh
        cX = lw
        cY = lh
    End With
    
End Sub
    
Public Function FileExist(sFileN As String) As Boolean
    Dim tmpRv As Long
    
    On Error Resume Next
    tmpRv = GetAttr(sFileN)
    If Err Then
        FileExist = False
    Else
        FileExist = True
    End If
End Function

Public Sub savPic()
    frmMain.picUndoTools.Height = frmMain.picMain.Height
    frmMain.picUndoTools.Width = frmMain.picMain.Width
    frmMain.picUndoTools.Picture = frmMain.picMain.Image
End Sub

Public Sub Filling(col As Long, ByVal FStyle As Long, x, y)
    Dim a As Long
    
    frmMain.picMain.FillStyle = FStyle
    frmMain.picMain.FillColor = tColors.lFCol
    a = ExtFloodFill(frmMain.picMain.hdc, x, y, col, 1)
    frmMain.picMain.FillStyle = 1
End Sub

Public Sub PrepFlip()
    With frmMain
        .picFlip.Height = .picMain.Height
        .picFlip.Width = .picMain.Width
        .picFlip.Picture = .picMain.Image
    End With
End Sub

Public Sub ReplaceColor()
    Dim tmp As Long
    
    FirstChoose = True
    For i = 0 To cX
        For j = 0 To cY
            tmp = GetPixel(frmMain.picMain.hdc, i, j)
            If tmp = tmpCol1 Then
                SetPixel frmMain.picMain.hdc, i, j, tmpCol2
            End If
        Next j
        frmMain.pg1.Value = i * 100 \ (cX - 1)
    Next i
    frmMain.pg1.Value = 0
End Sub

Public Sub Zoom(ZFactor As Double)
    
    '+25% => ZFactor = 1.25
    '+50% => ZFactor = 1.5
    '+75% => ZFactor = 1.75
    '-25% => ZFactor = 0.25
    '-50% => ZFactor = 0.50
    '-75% => ZFactor = 0.75
    'It could be that information get lost!
    'The Undo-Functions doesn't work well!
    'If anyone has a better solution email it to me, please.
    'SekKir@gmx.at
    Call Save
    frmMain.picZoom.Height = frmMain.picMain.Height * ZFactor
    frmMain.picZoom.Width = frmMain.picMain.Width * ZFactor
    frmMain.picZoom.PaintPicture frmMain.picMain.Picture, 0, 0, frmMain.picZoom.Width + ZFactor * 2, frmMain.picZoom.Height + ZFactor * 2, 0, 0, frmMain.picMain.Width, frmMain.picMain.Height, vbSrcCopy
    
    Set frmMain.picZoom.Picture = frmMain.picZoom.Image
    
    frmMain.picMain.Height = frmMain.picZoom.Height
    frmMain.picMain.Width = frmMain.picZoom.Width
    frmMain.picMain.PaintPicture frmMain.picZoom.Picture, 0, 0, , , , , , , vbSrcCopy
    
    Set frmMain.picMain.Picture = frmMain.picMain.Image
    Call ResizePicBoxes
    
End Sub
