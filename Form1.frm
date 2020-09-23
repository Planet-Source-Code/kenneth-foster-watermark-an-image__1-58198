VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Watermark an Image"
   ClientHeight    =   3330
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   3330
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   150
      MaxLength       =   2
      TabIndex        =   19
      Text            =   "8"
      Top             =   2925
      Width           =   345
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   165
      MaxLength       =   1
      TabIndex        =   17
      Text            =   "0"
      Top             =   2580
      Width           =   315
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1890
      Left            =   3960
      ScaleHeight     =   1890
      ScaleWidth      =   2250
      TabIndex        =   2
      Top             =   30
      Width           =   2250
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save as bmp"
      Enabled         =   0   'False
      Height          =   390
      Left            =   2670
      TabIndex        =   8
      Top             =   2100
      Width           =   1095
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load Image"
      Height          =   375
      Left            =   2670
      TabIndex        =   7
      Top             =   1650
      Width           =   1080
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   495
      Left            =   2655
      MultiLine       =   -1  'True
      TabIndex        =   6
      Text            =   "Form1.frx":0000
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Text"
      Enabled         =   0   'False
      Height          =   360
      Left            =   2745
      TabIndex        =   5
      Top             =   600
      Width           =   960
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Enabled         =   0   'False
      Height          =   330
      Left            =   2730
      TabIndex        =   4
      Top             =   120
      Width           =   990
   End
   Begin VB.VScrollBar sldrAlphaLevel 
      Height          =   1920
      Left            =   2385
      Max             =   255
      TabIndex        =   3
      Top             =   45
      Value           =   255
      Width           =   225
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   2295
      TabIndex        =   1
      Top             =   30
      Width           =   2295
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1860
      Left            =   3960
      ScaleHeight     =   1860
      ScaleWidth      =   2220
      TabIndex        =   0
      Top             =   45
      Width           =   2220
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5685
      TabIndex        =   22
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5400
      TabIndex        =   21
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Fontsize"
      Height          =   210
      Left            =   645
      TabIndex        =   20
      Top             =   2985
      Width           =   705
   End
   Begin VB.Label Label9 
      Caption         =   "How well a font color works depends on picture"
      Height          =   450
      Left            =   4095
      TabIndex        =   18
      Top             =   2925
      Width           =   2115
   End
   Begin VB.Label Label8 
      Caption         =   "Text Position(0 - 9)"
      Height          =   195
      Left            =   525
      TabIndex        =   16
      Top             =   2610
      Width           =   1410
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Text Color"
      Height          =   210
      Left            =   4650
      TabIndex        =   15
      Top             =   2415
      Width           =   870
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   5115
      TabIndex        =   14
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label5"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4830
      TabIndex        =   13
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4545
      TabIndex        =   12
      Top             =   2640
      Width           =   255
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4260
      TabIndex        =   11
      Top             =   2640
      Width           =   255
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   360
      Top             =   375
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Watermarked"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   345
      TabIndex        =   10
      Top             =   2055
      Width           =   1380
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Original"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4500
      TabIndex        =   9
      Top             =   2055
      Width           =   780
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   'Special note: AlphaBlending code is by Aaron DeRenad
   'and ModCmnDlg module is by Mr. Bobo
   
   Private Type AlphaOptions
   AlphaOption As Byte
   AlphaFlags As Byte
   SourceConstantAlpha As Byte
   AlphaFormat As Byte
End Type
Dim AO As AlphaOptions, newAO As Long
Dim AlphaIncrease As Boolean
Const AC_SRC_OVER = &H0
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Integer

Private Sub Form_Load()
   'set the graphic mode to 'persistent'
   Picture1.AutoRedraw = True
   Picture2.AutoRedraw = True
   
   Picture2.Width = Picture1.Width
   Picture2.Height = Picture1.Height
   
   With AO
      .AlphaOption = AC_SRC_OVER
      .AlphaFlags = 0
      .SourceConstantAlpha = 0
      .AlphaFormat = 0
   End With
   
End Sub

Public Sub CreateAlpha()
   
   'copy the AlphaOptions-structure to a Long
   RtlMoveMemory newAO, AO, 4
   'AlphaBlend the picture from Picture1 over the picture of Picture2
   Picture2.Picture = Picture3.Picture
   AlphaBlend Picture2.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture1.hdc, 0, 0, Picture1.ScaleWidth, Picture1.ScaleHeight, newAO
End Sub

Private Sub cmdClear_Click()
   Picture1.Cls
   Picture2.Cls
   Picture1.Picture = Picture3.Picture
End Sub

Private Sub cmdGenerate_Click()
   Dim x As Integer
   
   sldrAlphaLevel.Value = 255
   Picture2.FontSize = Int(Text3.Text)
   If Text2.Text = "0" Or Text2.Text = "" Then GoTo pt1
   For x = 1 To Int(Text2.Text)
       Picture2.Print
   Next x
   Picture2.Print Text1.Text
   BitBlt Picture1.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, vbSrcAnd
   cmdClear.Enabled = True
   cmdSave.Enabled = True
   Exit Sub
pt1:
   Picture2.Print Text1.Text
   BitBlt Picture1.hdc, 0, 0, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture2.hdc, 0, 0, vbSrcAnd
   cmdClear.Enabled = True
   cmdSave.Enabled = True
End Sub

Private Sub cmdLoad_Click()
   Picture1.Cls
   Picture2.Cls
   Picture3.Cls
   
   ShowOpen
   If cmndlg.FileName = "" Then Exit Sub
   Image1.Picture = LoadPicture(cmndlg.FileName)
   CreateThumb Picture1, Image1, 3000, 3000, True
   Picture2.Picture = Picture1.Image
   Picture3.Picture = Picture1.Image
   
   cmdGenerate.Enabled = True
End Sub

Private Sub cmdSave_Click()
   ShowSave
   Picture2.Picture = Picture2.Image
   If cmndlg.FileName = ".bmp" Or cmndlg.FileName = "" Then Exit Sub
   SavePicture Picture2.Picture, cmndlg.FileName & ".bmp"
End Sub

Private Sub Label11_Click()
   Picture2.ForeColor = Label11.BackColor
End Sub

Private Sub Label12_Click()
   Picture2.ForeColor = Label12.BackColor
End Sub

Private Sub Label3_Click()
   Picture2.ForeColor = Label3.BackColor
End Sub

Private Sub Label4_Click()
   Picture2.ForeColor = Label4.BackColor
End Sub

Private Sub Label5_Click()
   Picture2.ForeColor = Label5.BackColor
End Sub

Private Sub Label6_Click()
   Picture2.ForeColor = Label6.BackColor
End Sub

Private Sub sldrAlphaLevel_Change()
   AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
   Call CreateAlpha
End Sub

Private Sub sldrAlphaLevel_Click()
   AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
   Call CreateAlpha
End Sub

Private Sub sldrAlphaLevel_Scroll()
   AO.SourceConstantAlpha = CByte(sldrAlphaLevel.Value)
   Call CreateAlpha
End Sub

Public Sub CreateThumb(PicBox As Object, ByVal ActualPic As StdPicture, ByVal MaxHeight As Integer, ByVal MaxWidth As Integer, Center As Boolean, Optional ByVal PicTop As Integer, Optional ByVal PicLeft As Integer)
   'MaxHeight is max. image height allowed
   'MaxWidth is max. picture width allowed
   Dim NewH As Integer 'New Height
   Dim NewW As Integer 'New Width
   'set starting var.
   NewH = ActualPic.Height 'actual image height
   NewW = ActualPic.Width 'actual image width
   'do logic
   
   If NewH > MaxHeight Or NewW > MaxWidth Then 'picture is too large
   
   If NewH > NewW Then 'height is greater than width
   NewW = Fix((NewW / NewH) * MaxHeight) 'rescale height
   NewH = MaxHeight 'set max height
ElseIf NewW > NewH Then 'width is greater than height
   NewH = Fix((NewH / NewW) * MaxWidth) 'rescale width
   NewW = MaxHeight 'set max Width
   Debug.Print "Width>"
Else 'image is perfect square
   NewH = MaxHeight
   NewW = MaxWidth
End If
End If
'check if centered
If Center = True Then 'center picture
PicTop = (PicBox.Height / 2) - (NewH / 2)
PicLeft = (PicBox.Width / 2) - (NewW / 2)
Else 'if Optional variables are missing Then and center=false

If IsMissing(PicTop) = True Or IsMissing(PicLeft) = True Then
PicTop = 0 'Default top position
PicLeft = 0 'Default left position
End If
End If
'Draw newly scaled picture
With PicBox
.AutoRedraw = True 'set needed properties
.Cls 'clear picture box
.PaintPicture ActualPic, PicLeft, PicTop, NewW, NewH 'paint new picture size in picturebox
End With
End Sub
