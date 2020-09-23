VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8985
   ClientLeft      =   3720
   ClientTop       =   1125
   ClientWidth     =   12465
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   831
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picDest2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   105
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   809
      TabIndex        =   2
      Top             =   105
      Width           =   12135
      Begin VB.Shape shpRightBottom 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H00FF000A&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   11880
         Top             =   0
         Width           =   255
      End
      Begin VB.Shape shpLeftBottom 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   11880
         Top             =   8400
         Width           =   255
      End
      Begin VB.Shape shpLeftTop 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H0000FFFF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   8400
         Width           =   255
      End
      Begin VB.Shape shpRightTop 
         BorderColor     =   &H00FFFFFF&
         DrawMode        =   7  'Invert
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   0
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.PictureBox picSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   105
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   809
      TabIndex        =   0
      Top             =   105
      Width           =   12135
   End
   Begin VB.PictureBox picDest 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   8655
      Left            =   105
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   809
      TabIndex        =   1
      Top             =   105
      Width           =   12135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LeftTopX As Single, LeftTopY As Single
Dim RightTopX As Single, RightTopY As Single
Dim LeftBottomX As Single, LeftBottomY As Single
Dim RightBottomX As Single, RightBottomY As Single

Dim IsLeftTopOn As Boolean
Dim IsRightTopOn As Boolean
Dim IsLeftBottomOn As Boolean
Dim IsRightBottomOn As Boolean
Dim cLeftTopX, cLeftBottomX, cRightTopX, cRightBottomX
Dim cLeftTopY, cLeftBottomY, cRightTopY, cRightBottomY

Private Declare Function SetStretchBltMode Lib "gdi32" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Sub TransformVertical()
    On Error Resume Next
    
    picDest.Cls
    
    
    SetStretchBltMode picDest.hdc, 3
    For i = 0 To picDest.ScaleHeight
        cRightTopX = shpRightTop.Left - (i / (picSrc.ScaleWidth / (shpRightTop.Left - shpLeftTop.Left)))
        cRightBottomX = shpRightBottom.Left + (i / (picSrc.ScaleWidth / (picSrc.ScaleWidth - (shpRightBottom.Left - (shpLeftBottom.Left - picSrc.ScaleWidth)))))
        StretchBlt picDest.hdc, cRightTopX, i, cRightBottomX - cRightTopX, 1, picSrc.hdc, 0, i, picSrc.ScaleHeight, 1, vbSrcCopy
    Next
        
End Sub

Private Sub TransformHorizontal()
    On Error Resume Next
    
    picDest2.Cls
    
    SetStretchBltMode picDest2.hdc, 3
    For i = 0 To picDest2.ScaleWidth
        
        cRightTopY = shpRightTop.Top - (i / (picDest2.ScaleHeight / (shpRightTop.Top - shpRightBottom.Top)))
        cRightBottomY = shpLeftTop.Top + (i / (picDest2.ScaleHeight / (picDest2.ScaleHeight - (shpLeftTop.Top - (shpLeftBottom.Top - picDest2.ScaleHeight)))))
        
        StretchBlt picDest2.hdc, i, cRightTopY, 1, cRightBottomY - cRightTopY, picDest.hdc, i, 0, 1, picSrc.ScaleWidth, vbSrcCopy
    Next
    
End Sub

Private Sub Form_Load()
    
    picSrc.PaintPicture picSrc.Picture, 0, 0, picSrc.ScaleWidth, picSrc.ScaleHeight
    
    picDest2_MouseMove 1, 0, shpLeftBottom.Left + 1, shpLeftTop.Top + 1
    picDest2_MouseDown 1, 0, shpLeftBottom.Left + 1, shpLeftTop.Top + 1
    
End Sub

Private Sub picDest_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.Caption = picDest.Point(x, y)
End Sub

Private Sub picDest2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
    
        If IsInObject(shpLeftTop, x, y) = True Then
            LeftTopX = x - shpLeftTop.Left
            LeftTopY = y - shpLeftTop.Top
            IsLeftTopOn = True
        End If
        
        If IsInObject(shpRightTop, x, y) = True Then
            RightTopX = x - shpRightTop.Left
            RightTopY = y - shpRightTop.Top
            IsRightTopOn = True
        End If
        
        If IsInObject(shpLeftBottom, x, y) = True Then
            LeftBottomX = x - shpLeftBottom.Left
            LeftBottomY = y - shpLeftBottom.Top
            IsLeftBottomOn = True
        End If
        
        If IsInObject(shpRightBottom, x, y) = True Then
            RightBottomX = x - shpRightBottom.Left
            RightBottomY = y - shpRightBottom.Top
            IsRightBottomOn = True
        End If
        
    End If
End Sub

Private Sub picDest2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If (IsLeftTopOn) Or (IsRightTopOn) Or (IsLeftBottomOn) Or (IsRightBottomOn) Then
        TransformHorizontal
        TransformVertical
    End If
    If Button = 1 Then
        If IsLeftTopOn = True Then
            shpLeftTop.Left = x - LeftTopX
            shpLeftTop.Top = y - LeftTopY
        End If
        
        If IsRightTopOn = True Then
            shpRightTop.Left = x - RightTopX
            shpRightTop.Top = y - RightTopY
        End If
        
        If IsLeftBottomOn = True Then
            shpLeftBottom.Left = x - LeftBottomX
            shpLeftBottom.Top = y - LeftBottomY
        End If
        
        If IsRightBottomOn = True Then
            shpRightBottom.Left = x - RightBottomX
            shpRightBottom.Top = y - RightBottomY
        End If
        
    End If
End Sub

Private Sub picDest2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    IsLeftTopOn = False
    IsRightTopOn = False
    IsLeftBottomOn = False
    IsRightBottomOn = False
End Sub
Private Function IsInObject(Object As Object, x, y) As Boolean
    If ((y > Object.Top) And (y < Object.Height + Object.Top)) And _
    ((x > Object.Left) And (x < Object.Width + Object.Left)) Then
        IsInObject = True
    Else
        IsInObject = False
    End If
End Function

Private Sub HideShowShapes(TrueOrFalse As Boolean)
    shpLeftBottom.Visible = TrueOrFalse
    shpLeftTop.Visible = TrueOrFalse
    shpRightBottom.Visible = TrueOrFalse
    shpRightTop.Visible = TrueOrFalse
End Sub
