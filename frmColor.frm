VERSION 5.00
Begin VB.Form frmColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Advance Color Picker"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11670
   Icon            =   "frmColor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pctSmallCube 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6330
      ScaleHeight     =   405
      ScaleWidth      =   405
      TabIndex        =   42
      Top             =   1365
      Width           =   435
   End
   Begin VB.HScrollBar HSZ 
      Height          =   255
      LargeChange     =   15
      Left            =   7320
      TabIndex        =   38
      Top             =   1905
      Width           =   1350
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   7320
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   1200
      Width           =   1365
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmColor.frx":0442
      Left            =   7320
      List            =   "frmColor.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   33
      Top             =   690
      Width           =   1365
   End
   Begin VB.PictureBox pctCube 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2205
      Left            =   8760
      MouseIcon       =   "frmColor.frx":048A
      MousePointer    =   99  'Custom
      ScaleHeight     =   2175
      ScaleWidth      =   2175
      TabIndex        =   31
      Top             =   225
      Width           =   2205
      Begin VB.Image imgSelColor 
         Height          =   480
         Left            =   315
         Picture         =   "frmColor.frx":05DC
         Top             =   795
         Visible         =   0   'False
         Width           =   480
      End
   End
   Begin VB.TextBox txtColorHEX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6090
      TabIndex        =   28
      Text            =   "DDDDDD"
      Top             =   2175
      Width           =   1065
   End
   Begin VB.Frame frmHSL 
      Caption         =   "HSL"
      Height          =   1170
      Left            =   105
      TabIndex        =   12
      Top             =   1320
      Width           =   5550
      Begin VB.TextBox txtLum 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   795
         Width           =   705
      End
      Begin VB.TextBox txtSat 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   495
         Width           =   705
      End
      Begin VB.TextBox txtHue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   195
         Width           =   705
      End
      Begin VB.PictureBox pctLum 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   18
         Top             =   810
         Width           =   750
      End
      Begin VB.PictureBox pctSat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   17
         Top             =   510
         Width           =   750
      End
      Begin VB.PictureBox pctHue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   16
         Top             =   195
         Width           =   750
      End
      Begin VB.HScrollBar HSLum 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   240
         TabIndex        =   15
         Top             =   825
         Width           =   2895
      End
      Begin VB.HScrollBar HSSat 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   240
         TabIndex        =   14
         Top             =   510
         Width           =   2895
      End
      Begin VB.HScrollBar HSHue 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   239
         TabIndex        =   13
         Top             =   210
         Width           =   2895
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Luminance:"
         Height          =   255
         Left            =   30
         TabIndex        =   27
         Top             =   825
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Saturation:"
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   510
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Hue:"
         Height          =   255
         Left            =   30
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame frmRGB 
      Caption         =   "RGB"
      Height          =   1170
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   5550
      Begin VB.HScrollBar HSRed 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   255
         TabIndex        =   11
         Top             =   210
         Width           =   2895
      End
      Begin VB.HScrollBar HSGreen 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   255
         TabIndex        =   10
         Top             =   510
         Width           =   2895
      End
      Begin VB.HScrollBar HSBlue 
         Height          =   240
         LargeChange     =   15
         Left            =   915
         Max             =   255
         TabIndex        =   9
         Top             =   825
         Width           =   2895
      End
      Begin VB.PictureBox pctRed 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   8
         Top             =   195
         Width           =   750
      End
      Begin VB.PictureBox pctGreen 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   7
         Top             =   510
         Width           =   750
      End
      Begin VB.PictureBox pctBlue 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3900
         ScaleHeight     =   225
         ScaleWidth      =   720
         TabIndex        =   6
         Top             =   810
         Width           =   750
      End
      Begin VB.TextBox txtRed 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   195
         Width           =   705
      End
      Begin VB.TextBox txtGreen 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   495
         Width           =   705
      End
      Begin VB.TextBox txtBlue 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4725
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   795
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Blue:"
         Height          =   255
         Left            =   30
         TabIndex        =   24
         Top             =   825
         Width           =   870
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Red:"
         Height          =   255
         Left            =   30
         TabIndex        =   23
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Green:"
         Height          =   255
         Left            =   30
         TabIndex        =   22
         Top             =   525
         Width           =   870
      End
   End
   Begin VB.TextBox txtColor 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6090
      TabIndex        =   1
      Text            =   "16777215"
      Top             =   1875
      Width           =   1065
   End
   Begin VB.PictureBox pctColor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1065
      Left            =   5715
      ScaleHeight     =   1035
      ScaleWidth      =   1410
      TabIndex        =   0
      Top             =   225
      Width           =   1440
   End
   Begin VB.Label lblToggleColorSquare 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   ">>"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   6300
      TabIndex        =   32
      ToolTipText     =   "Toggle Color Square"
      Top             =   1335
      Width           =   855
   End
   Begin VB.Label lblShowColor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected Color »"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7335
      TabIndex        =   41
      ToolTipText     =   "Pick Color from Color Square"
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label lblZVal 
      Alignment       =   2  'Center
      Caption         =   "Label11"
      Height          =   180
      Left            =   7335
      TabIndex        =   40
      Top             =   1710
      Width           =   1350
   End
   Begin VB.Label lblZ 
      Caption         =   "Z Axis:"
      Height          =   210
      Left            =   7320
      TabIndex        =   39
      Top             =   1545
      Width           =   1335
   End
   Begin VB.Label lblUpdateSquare 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Update Square »"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7335
      TabIndex        =   37
      ToolTipText     =   "Pick Color from Color Square"
      Top             =   2190
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Y Axis:"
      Height          =   210
      Left            =   7320
      TabIndex        =   36
      Top             =   1005
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "X Axis:"
      Height          =   210
      Left            =   7335
      TabIndex        =   34
      Top             =   495
      Width           =   1335
   End
   Begin VB.Line Line3 
      X1              =   11145
      X2              =   11145
      Y1              =   285
      Y2              =   2460
   End
   Begin VB.Line Line1 
      Visible         =   0   'False
      X1              =   7350
      X2              =   7350
      Y1              =   225
      Y2              =   2460
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "#"
      Height          =   195
      Left            =   5760
      TabIndex        =   30
      Top             =   1935
      Width           =   270
   End
   Begin VB.Label Label7 
      Caption         =   "Hex"
      Height          =   225
      Left            =   5730
      TabIndex        =   29
      Top             =   2190
      Width           =   300
   End
   Begin VB.Image imgDummy 
      Height          =   480
      Left            =   6930
      Picture         =   "frmColor.frx":072E
      Top             =   495
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   510
      Left            =   5730
      MouseIcon       =   "frmColor.frx":0880
      Picture         =   "frmColor.frx":09D2
      ToolTipText     =   "Drag the bullseye to pick color"
      Top             =   1335
      Width           =   510
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   510
      Left            =   5730
      Top             =   1335
      Width           =   510
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000010&
      BorderWidth     =   3
      Visible         =   0   'False
      X1              =   7335
      X2              =   7335
      Y1              =   210
      Y2              =   2475
   End
End
Attribute VB_Name = "frmColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DontGetSetColor As Boolean
Dim ColorSquareDrawn As Boolean
Dim ColorSquareXYZ As String, ColorSquareZVal As Integer

Private Sub RefreshSmallCube()
    If pctSmallCube.Tag <> ColorSquareXYZ & ColorSquareZVal Then
        DrawColorSquare pctSmallCube, ColorSquareXYZ, ColorSquareZVal
        pctSmallCube.Tag = ColorSquareXYZ & ColorSquareZVal
    End If
End Sub

Private Sub RefreshCube()
    If pctCube.Tag <> ColorSquareXYZ & ColorSquareZVal Then
        imgSelColor.Visible = False
        DrawColorSquare pctCube, ColorSquareXYZ, ColorSquareZVal
        pctCube.Tag = ColorSquareXYZ & ColorSquareZVal
    End If
End Sub
Private Sub Combo1_Click()
    Combo2.Clear
    Select Case Combo1.Text
        Case "Red"
            Combo2.AddItem "Green"
            Combo2.AddItem "Blue"
        Case "Green"
            Combo2.AddItem "Red"
            Combo2.AddItem "Blue"
        Case "Blue"
            Combo2.AddItem "Red"
            Combo2.AddItem "Green"
        Case "Hue"
            Combo2.AddItem "Saturation"
            Combo2.AddItem "Luminance"
        Case "Saturation"
            Combo2.AddItem "Hue"
            Combo2.AddItem "Luminance"
        Case "Luminance"
            Combo2.AddItem "Hue"
            Combo2.AddItem "Saturation"
    End Select
End Sub

Private Sub Combo2_Click()
    Select Case Left(Combo1.Text, 1) & Left(Combo2.Text, 1)
        Case "RG": ColorSquareXYZ = "RGB"
        Case "RB": ColorSquareXYZ = "RBG"
        Case "GR": ColorSquareXYZ = "GRB"
        Case "GB": ColorSquareXYZ = "GBR"
        Case "BR": ColorSquareXYZ = "BRG"
        Case "BG": ColorSquareXYZ = "BGR"
    
        Case "HS": ColorSquareXYZ = "HSL"
        Case "HL": ColorSquareXYZ = "HLS"
        Case "SH": ColorSquareXYZ = "SHL"
        Case "SL": ColorSquareXYZ = "SLH"
        Case "LH": ColorSquareXYZ = "LHS"
        Case "LS": ColorSquareXYZ = "LSH"
    End Select
    
    Select Case Right(ColorSquareXYZ, 1)
        Case "R": lblZ.Caption = "Z Axis: Red": HSZ.Max = 255
        Case "G": lblZ.Caption = "Z Axis: Green": HSZ.Max = 255
        Case "B": lblZ.Caption = "Z Axis: Blue": HSZ.Max = 255
        Case "H": lblZ.Caption = "Z Axis: Hue": HSZ.Max = HueMAX
        Case "S": lblZ.Caption = "Z Axis: Saturation": HSZ.Max = SatMAX
        Case "L": lblZ.Caption = "Z Axis: Luminance": HSZ.Max = LumMAX
    End Select
    'ColorSquareZVal = HSZ.Max / 2
    HSZ.Value = HSZ.Max / 2
End Sub

Private Sub Form_Load()
pctSmallCube.ToolTipText = lblToggleColorSquare.ToolTipText
ColorSquareXYZ = "HSL"
ColorSquareZVal = 120
RefreshSmallCube
HSHue.Max = HueMAX
HSSat.Max = SatMAX
HSLum.Max = LumMAX
Me.Width = Line1.X1
GetColorFromScroll_RGB
End Sub

Private Sub GetColorFromScroll_RGB(Optional DontSetText As Boolean)
Dim HSLis As HSL
If DontGetSetColor Then Exit Sub
imgSelColor.Visible = False
DontGetSetColor = True
pctRed.BackColor = RGB(HSRed.Value, 0, 0)
If Not DontSetText Then
    txtRed.Text = HSRed.Value
    txtRed.ForeColor = vbBlack
End If
pctGreen.BackColor = RGB(0, HSGreen.Value, 0)
If Not DontSetText Then
    txtGreen.Text = HSGreen.Value
    txtGreen.ForeColor = vbBlack
End If
pctBlue.BackColor = RGB(0, 0, HSBlue.Value)
If Not DontSetText Then
    txtBlue.Text = HSBlue.Value
    txtBlue.ForeColor = vbBlack
End If
pctColor.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
If Not DontSetText Then
    txtColor.Text = pctColor.BackColor
    txtColorHEX.Text = Hex(txtColor.Text)
    txtColor.ForeColor = vbBlack
    txtColorHEX.ForeColor = vbBlack
End If

HSLis = RGBtoHSL(HSRed.Value, HSGreen.Value, HSBlue.Value)
HSHue.Value = HSLis.Hue
If Not DontSetText Then
    txtHue.Text = HSHue.Value
    txtHue.ForeColor = vbBlack
End If
pctHue.BackColor = HSL(HSHue.Value, 240, 120)
HSSat.Value = HSLis.Saturation
If Not DontSetText Then
    txtSat.Text = HSSat.Value
    txtSat.ForeColor = vbBlack
End If
pctSat.BackColor = HSL(HSHue.Value, HSSat.Value, 120)
HSLum.Value = HSLis.Luminance
If Not DontSetText Then
    txtLum.Text = HSLum.Value
    txtLum.ForeColor = vbBlack
End If
pctLum.BackColor = HSL(HSHue.Value, HSSat.Value, HSLum.Value)
DontGetSetColor = False
End Sub

Private Sub GetColorFromScroll_HSL(Optional DontSetText As Boolean)
Dim RGBis As RGB
If DontGetSetColor Then Exit Sub
imgSelColor.Visible = False
DontGetSetColor = True
If Not DontSetText Then
    txtHue.Text = HSHue.Value
    txtHue.ForeColor = vbBlack
End If
pctHue.BackColor = HSL(HSHue.Value, HueMAX, Int(LumMAX / 2))
If Not DontSetText Then
    txtSat.Text = HSSat.Value
    txtSat.ForeColor = vbBlack
End If
pctSat.BackColor = HSL(HSHue.Value, HSSat.Value, Int(LumMAX / 2))
If Not DontSetText Then
    txtLum.Text = HSLum.Value
    txtLum.ForeColor = vbBlack
End If
pctLum.BackColor = HSL(HSHue.Value, HueMAX, HSLum.Value)

RGBis = HSLtoRGB(HSHue.Value, HSSat.Value, HSLum.Value)
HSRed.Value = RGBis.Red
pctRed.BackColor = RGB(HSRed.Value, 0, 0)
If Not DontSetText Then
    txtRed.Text = HSRed.Value
    txtRed.ForeColor = vbBlack
End If
HSGreen.Value = RGBis.Green
pctGreen.BackColor = RGB(0, HSGreen.Value, 0)
If Not DontSetText Then
    txtGreen.Text = HSGreen.Value
    txtGreen.ForeColor = vbBlack
End If
HSBlue.Value = RGBis.Blue
pctBlue.BackColor = RGB(0, 0, HSBlue.Value)
If Not DontSetText Then
    txtBlue.Text = HSBlue.Value
    txtBlue.ForeColor = vbBlack
End If
pctColor.BackColor = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
If Not DontSetText Then
    txtColor.Text = pctColor.BackColor
    txtColorHEX.Text = Hex(txtColor.Text)
    txtColor.ForeColor = vbBlack
    txtColorHEX.ForeColor = vbBlack
End If

DontGetSetColor = False
End Sub


Private Sub HSBlue_Change()
GetColorFromScroll_RGB
End Sub

Private Sub HSBlue_Scroll()
GetColorFromScroll_RGB
End Sub

Private Sub HSGreen_Change()
GetColorFromScroll_RGB
End Sub

Private Sub HSGreen_Scroll()
GetColorFromScroll_RGB
End Sub

Private Sub HSHue_Change()
GetColorFromScroll_HSL
End Sub

Private Sub HSHue_Scroll()
GetColorFromScroll_HSL
End Sub

Private Sub HSLum_Change()
GetColorFromScroll_HSL
End Sub

Private Sub HSLum_Scroll()
GetColorFromScroll_HSL
End Sub

Private Sub HSRed_Change()
GetColorFromScroll_RGB
End Sub

Private Sub HSRed_Scroll()
GetColorFromScroll_RGB
End Sub

Private Sub HSSat_Change()
GetColorFromScroll_HSL
End Sub

Private Sub HSSat_Scroll()
GetColorFromScroll_HSL
End Sub

Private Sub HSZ_Change()
lblZVal.Caption = HSZ.Value
ColorSquareZVal = HSZ.Value
RefreshSmallCube
Me.MousePointer = 11 'Hourglass
DoEvents
RefreshCube
Me.MousePointer = 0 'Default
End Sub

Private Sub HSZ_Scroll()
lblZVal.Caption = HSZ.Value
ColorSquareZVal = HSZ.Value
RefreshSmallCube
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
    If Image1.MousePointer = 0 Then
        Image1.MousePointer = 99
        Set Image1.Picture = Nothing
    End If
    Image1.Refresh
    Shape1.FillColor = GetColorAtCursor
    'pctColor.BackColor = GetColorAtCursor
    'txtColor.Text = GetColorAtCursor
Else
    If Image1.MousePointer = 99 Then
        Image1.MousePointer = 0
        Set Image1.Picture = imgDummy.Picture
    End If
End If
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    If Image1.MousePointer = 99 Then
        Image1.MousePointer = 0
        Set Image1.Picture = imgDummy.Picture
        Shape1.FillColor = vbWhite
        txtColor.Text = GetColorAtCursor
    End If
End If
End Sub

Private Sub ToggleColorSquare()
If Me.Width = Line1.X1 Then
    Me.Width = Line3.X1
    If Not ColorSquareDrawn Then
        Combo1.ListIndex = 3
        Combo2.ListIndex = 0
        ColorSquareDrawn = True
    End If
    lblToggleColorSquare.Caption = "<<"
Else
    Me.Width = Line1.X1
    lblToggleColorSquare.Caption = ">>"
End If
End Sub

Private Sub imgSelColor_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctCube_MouseDown Button, Shift, X + imgSelColor.Left, Y + imgSelColor.Top
End Sub

Private Sub imgSelColor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    pctCube_MouseMove Button, Shift, X + imgSelColor.Left, Y + imgSelColor.Top
End Sub

Private Sub lblShowColor_Click()
    ShowColorInSquare
End Sub

Private Sub lblToggleColorSquare_Click()
ToggleColorSquare
End Sub

Private Sub lblToggleColorSquare_DblClick()
lblToggleColorSquare_Click
End Sub

Private Sub lblUpdateSquare_Click()
Me.MousePointer = 11 'Hourglass
DoEvents
RefreshCube
Me.MousePointer = 0 'Default
End Sub

Private Sub pctCube_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    txtColor.Text = pctCube.Point(X, Y)
    imgSelColor.Visible = False
End If
End Sub

Private Sub pctCube_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
    txtColor.Text = pctCube.Point(X, Y)
End If
End Sub

Private Sub ShowColorInSquare()
'ColorSquareXYZ ColorSquareZVal
Dim XVal As Integer, YVal As Integer, ZVal As Integer
    If ColorSquareDrawn Then
        With pctCube
            Select Case Left(ColorSquareXYZ, 2)
                Case "RG": XVal = .Width / 255 * HSRed.Value: YVal = .Height / 255 * HSGreen.Value: ZVal = HSBlue.Value
                Case "RB": XVal = .Width / 255 * HSRed.Value: ZVal = HSGreen.Value: YVal = .Height / 255 * HSBlue.Value
                Case "GR": YVal = .Height / 255 * HSRed.Value: XVal = .Width / 255 * HSGreen.Value: ZVal = HSBlue.Value
                Case "GB": ZVal = HSRed.Value: XVal = .Width / 255 * HSGreen.Value: YVal = .Height / 255 * HSBlue.Value
                Case "BR": YVal = .Height / 255 * HSRed.Value: ZVal = HSGreen.Value: XVal = .Width / 255 * HSBlue.Value
                Case "BG": ZVal = HSRed.Value: YVal = .Height / 255 * HSGreen.Value: XVal = .Width / 255 * HSBlue.Value
            
                Case "HS": XVal = .Width / HueMAX * HSHue.Value: YVal = .Height / SatMAX * HSSat.Value: ZVal = HSLum.Value
                Case "HL": XVal = .Width / HueMAX * HSHue.Value: ZVal = HSSat.Value: YVal = .Height / LumMAX * HSLum.Value
                Case "SH": YVal = .Height / HueMAX * HSHue.Value: XVal = .Width / SatMAX * HSSat.Value: ZVal = HSLum.Value
                Case "SL": ZVal = HSHue.Value: XVal = .Width / SatMAX * HSSat.Value: YVal = .Height / LumMAX * HSLum.Value
                Case "LH": YVal = .Height / HueMAX * HSHue.Value: ZVal = HSSat.Value: XVal = .Width / LumMAX * HSLum.Value
                Case "LS": ZVal = HSHue.Value: YVal = .Height / SatMAX * HSSat.Value: XVal = .Width / LumMAX * HSLum.Value
            End Select
            If HSZ.Value <> ZVal Then HSZ.Value = ZVal
            imgSelColor.Move XVal - imgSelColor.Width / 2 + 10, YVal - imgSelColor.Height / 2 + 10
            imgSelColor.Visible = True
        End With
    End If
End Sub

Private Sub pctCube_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
With imgSelColor
    .Move X - .Width / 2 + 15, Y - .Height / 2 + 15
    .Visible = True
End With
End Sub

Private Sub pctSmallCube_Click()
lblToggleColorSquare_Click
End Sub

Private Sub pctSmallCube_DblClick()
lblToggleColorSquare_Click
End Sub

Private Sub txtBlue_Change()
GetColorFromText "B"
End Sub

Private Sub txtColor_Change()
GetColorFromText "C"
End Sub

Private Sub txtColorHEX_Change()
GetColorFromText "CH"
End Sub

Private Sub txtGreen_Change()
GetColorFromText "G"
End Sub

Private Sub txtHue_Change()
GetColorFromText "H"
End Sub

Private Sub txtLum_Change()
GetColorFromText "L"
End Sub

Private Sub txtRed_Change()
GetColorFromText "R"
End Sub

Private Sub GetColorFromText(ChangedText As String)
Dim C As Long, HSLis As HSL
On Error GoTo ErrorHere
    If DontGetSetColor Then Exit Sub
    imgSelColor.Visible = False
    If ChangedText = "C" And CLng(txtColor.Text) > 16777215 Then
        txtColor.ForeColor = vbRed
        Exit Sub
    ElseIf ChangedText = "C" And CLng(txtColor.Text) <> txtColor.Text Then
        txtColor.ForeColor = vbRed
        Exit Sub
    ElseIf ChangedText = "CH" And CLng("&H" & txtColorHEX.Text) > 16777215 Then
        txtColorHEX.ForeColor = vbRed
        Exit Sub
    ElseIf ChangedText = "CH" And CLng("&H" & txtColorHEX.Text) <> CSng("&H" & txtColorHEX.Text) Then
        txtColorHEX.ForeColor = vbRed
        Exit Sub
    End If
    DontGetSetColor = True
    Select Case ChangedText
        Case "R"
            If CInt(txtRed.Text) <> txtRed.Text Then
                txtRed.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSRed.Value = txtRed.Text
            txtColor.Text = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "G"
            If CInt(txtGreen.Text) <> txtGreen.Text Then
                txtGreen.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSGreen.Value = txtGreen.Text
            txtColor.Text = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "B"
            If CInt(txtBlue.Text) <> txtBlue.Text Then
                txtBlue.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSBlue.Value = txtBlue.Text
            txtColor.Text = RGB(HSRed.Value, HSGreen.Value, HSBlue.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "H"
            If CInt(txtHue.Text) <> txtHue.Text Then
                txtHue.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSHue.Value = txtHue.Text
            txtColor.Text = HSL(HSHue.Value, HSSat.Value, HSLum.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "S"
            If CInt(txtSat.Text) <> txtSat.Text Then
                txtSat.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSSat.Value = txtSat.Text
            txtColor.Text = HSL(HSHue.Value, HSSat.Value, HSLum.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "L"
            If CInt(txtLum.Text) <> txtLum.Text Then
                txtLum.ForeColor = vbRed
                DontGetSetColor = False
                Exit Sub
            End If
            HSLum.Value = txtLum.Text
            txtColor.Text = HSL(HSHue.Value, HSSat.Value, HSLum.Value)
            txtColorHEX.Text = Hex(txtColor.Text)
        Case "C", "CH"
            If ChangedText = "C" Then
                C = txtColor.Text
                txtColorHEX.Text = Hex(C)
            Else
                C = CLng("&H" & txtColorHEX.Text)
                txtColor.Text = C
            End If
            HSRed.Value = C Mod 256
            txtRed.Text = HSRed.Value
            C = Int(C / 256)
            HSGreen.Value = C Mod 256
            txtGreen.Text = HSGreen.Value
            HSBlue.Value = Int(C / 256)
            txtBlue.Text = HSBlue.Value
            HSLis = RGBtoHSL(HSRed.Value, HSGreen.Value, HSBlue.Value)
            HSHue.Value = HSLis.Hue
            txtHue.Text = HSLis.Hue
            HSSat.Value = HSLis.Saturation
            txtSat.Text = HSLis.Saturation
            HSLum.Value = HSLis.Luminance
            txtLum.Text = HSLis.Luminance
    End Select
    txtColor.ForeColor = vbBlack
    txtColorHEX.ForeColor = vbBlack
    DontGetSetColor = False
    Select Case ChangedText
        Case "R", "G", "B", "C", "CH"
            txtHue.ForeColor = vbBlack
            txtSat.ForeColor = vbBlack
            txtLum.ForeColor = vbBlack
            GetColorFromScroll_RGB True
            DontGetSetColor = True
            txtHue.Text = HSHue.Value
            txtSat.Text = HSSat.Value
            txtLum.Text = HSLum.Value
            DontGetSetColor = False
        Case "H", "S", "L"
            txtRed.ForeColor = vbBlack
            txtGreen.ForeColor = vbBlack
            txtBlue.ForeColor = vbBlack
            GetColorFromScroll_HSL True
            DontGetSetColor = True
            txtRed.Text = HSRed.Value
            txtGreen.Text = HSGreen.Value
            txtBlue.Text = HSBlue.Value
            DontGetSetColor = False
    End Select
    Select Case ChangedText
        Case "R"
            txtRed.ForeColor = vbBlack
        Case "G"
            txtGreen.ForeColor = vbBlack
        Case "B"
            txtBlue.ForeColor = vbBlack
        Case "C"
            txtColor.ForeColor = vbBlack
        Case "H"
            txtHue.ForeColor = vbBlack
        Case "S"
            txtSat.ForeColor = vbBlack
        Case "L"
            txtLum.ForeColor = vbBlack
    End Select
    Exit Sub
ErrorHere:
    Select Case ChangedText
        Case "R"
            txtRed.ForeColor = vbRed
        Case "G"
            txtGreen.ForeColor = vbRed
        Case "B"
            txtBlue.ForeColor = vbRed
        Case "H"
            txtHue.ForeColor = vbRed
        Case "S"
            txtSat.ForeColor = vbRed
        Case "L"
            txtLum.ForeColor = vbRed
        Case "C"
            txtColor.ForeColor = vbRed
        Case "CH"
            txtColorHEX.ForeColor = vbRed
    End Select
    DontGetSetColor = False
End Sub

Private Sub txtSat_Change()
GetColorFromText "S"
End Sub
