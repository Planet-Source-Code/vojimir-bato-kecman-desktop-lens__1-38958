VERSION 5.00
Begin VB.Form FastLens 
   AutoRedraw      =   -1  'True
   ClientHeight    =   4455
   ClientLeft      =   3810
   ClientTop       =   3150
   ClientWidth     =   7320
   ControlBox      =   0   'False
   ForeColor       =   &H00FFFFFF&
   Icon            =   "Lens.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   297
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   488
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Pic 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4455
      Left            =   0
      ScaleHeight     =   297
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1024
      TabIndex        =   0
      Top             =   0
      Width           =   15360
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   3960
      Top             =   3120
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3840
      Top             =   3480
   End
End
Attribute VB_Name = "FastLens"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Simple Lens Demo
' by scythe scythe@cablenet.de
Private HSSIZE As Long
Private HSMAGN As Single
Private hScreenWidth As Long
Private hScreenHeight As Long

' Compile for real speed

'This demo uses a precalculated lens
'The array LookUp hold the difference between
'the point to set and the point to read
'We dont need to calculate the whole thing every cycle
'All we need is to draw the lens

Option Explicit

'To copy our pic real fast
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Use DIB for fast GFX
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

Private Type BITMAPINFOHEADER
 biSize           As Long
 biWidth          As Long
 biHeight         As Long
 biPlanes         As Integer
 biBitCount       As Integer
 biCompression    As Long
 biSizeImage      As Long
 biXPelsPerMeter  As Long
 biYPelsPerMeter  As Long
 biClrUsed        As Long
 biClrImportant   As Long
End Type

Private Type BITMAPINFO
 bmiHeader As BITMAPINFOHEADER
End Type

Private Const DIB_RGB_COLORS As Long = 0

Private Type PointApi
 x As Long
 y As Long
End Type

Dim LookUp() As PointApi      'Table for precalculatet Lens

Dim PicNew()  As RGBQUAD      'Hold our New Picture
Dim PicOrg()  As RGBQUAD      'Hold our Original Picute
Dim Binfo     As BITMAPINFO   'The GetDIBits API needs some Infos
Dim OrgLng    As Long         'Holds the Lenght of the Picture
Dim Drawing   As Boolean      'Is the program at work
Dim MoveX     As Long         'Holds X in Automove Mode
Dim MoveY     As Long         'Holds Y in Automove Mode
Dim DirX      As Byte         'Holds DirectionX in Automove Mode
Dim DirY      As Byte         'Holds DirectionY in Automove Mode

Private Sub Form_KeyPress(KeyAscii As Integer)
 If KeyAscii = 27 Then
  Timer1.Enabled = False
  Unload Me
  End
 End If
End Sub

Private Sub Form_Load()
    Dim hScreenDC As Long
    Dim hCompatibleDC As Long
    
    Dim hBitmap As Long
    Dim hBitmapOld As Long
    
    Dim hRetValue As Long
    
    HSSIZE = 200
    HSMAGN = 30
    
    'Grabb an image of the screen.
    hScreenDC = CreateDC("DISPLAY", "", "", 0&)
    hCompatibleDC = CreateCompatibleDC(hScreenDC)
    'Get screen size.
    hScreenWidth = GetDeviceCaps(hScreenDC, 8)
    hScreenHeight = GetDeviceCaps(hScreenDC, 10)
    'Copy it onto a bitmap and use for your own pleasure or amusement.
    hBitmap = CreateCompatibleBitmap(hScreenDC, hScreenWidth, hScreenHeight)
    hBitmapOld = SelectObject(hCompatibleDC, hBitmap)
    hRetValue = BitBlt(hCompatibleDC, 0, 0, hScreenWidth, hScreenHeight, hScreenDC, 0, 0, 13369376)
    hBitmap = SelectObject(hCompatibleDC, hBitmapOld)
    
    'Create a buffer that holds our picture
    ReDim PicNew(0 To hScreenWidth - 1, 0 To hScreenHeight - 1)
    ReDim PicOrg(0 To hScreenWidth - 1, 0 To hScreenHeight - 1)

    'Get the Picturesize in Memory for CopyMemory
    'X*Y*4 (4 for the 4 Bytes of RGBQUAD)
    OrgLng = (UBound(PicOrg, 1) + 1) * (UBound(PicOrg, 2) + 1) * 4

    'Set the infos for our apicall
    With Binfo.bmiHeader
        .biSize = 40
        .biWidth = hScreenWidth
        .biHeight = hScreenHeight
        .biPlanes = 1
        .biBitCount = 32
        .biCompression = 0
        .biClrUsed = 0
        .biClrImportant = 0
        .biSizeImage = hScreenWidth * hScreenHeight
    End With

    'Now get the Screen
    GetDIBits hScreenDC, hBitmap, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS
    'hScreenDC.CurrentX = 100
    'hScreenDC.CurrentY = 50
    'hScreenDC.Print "Please compile to get full SPEED"

    'Clean up the mess
    hRetValue = DeleteDC(hScreenDC)
    hRetValue = DeleteDC(hCompatibleDC)
    DeleteObject (hBitmapOld)
    DeleteObject (hBitmap)

    'Copy from Screen to Pic
    ShowPic

    'Calculate our Lens
    CreateLens LookUp, HSSIZE, HSMAGN '200,30

    'set direction and position for Automove
    DirX = 1
    DirY = 1
    MoveX = hScreenWidth / 2
    MoveY = hScreenHeight / 2

End Sub

Private Sub DrawLens(ByVal SourceX As Long, ByVal SourceY As Long)
 Dim x As Long
 Dim y As Long
 Dim StartX As Long
 Dim StartY As Long
 Dim EndX As Long
 Dim EndY As Long

 'Tell the program that we draw
 Drawing = True

 'Center our lens
 SourceX = SourceX - UBound(LookUp) / 2
 SourceY = SourceY - UBound(LookUp) / 2

 'Ok if we move the lens out of the Picture
 'Draw only the vissible part
 StartX = 2
 EndX = UBound(LookUp, 1) - 1
 If SourceX < 0 Then
  StartX = Abs(SourceX)
  ElseIf SourceX > hScreenWidth Then
  EndX = EndX + SourceX - hScreenWidth
 End If
 StartY = 1
 EndY = UBound(LookUp, 1) - 1
 If SourceY < 0 Then
  StartY = Abs(SourceY)
  ElseIf SourceY > hScreenHeight Then
  EndY = EndY + SourceY - hScreenHeight
 End If


On Error Resume Next
'Get the picture to paint on
CopyMemory PicNew(0, 0), PicOrg(0, 0), OrgLng

For x = StartX To EndX
 For y = StartY To EndY
  'Now we use our Array
  'Set a new point to x,y
  'Get the position on the original picture from our Array
  PicNew(x + SourceX - 1, y + SourceY) = PicOrg(LookUp(x, y).x + SourceX, LookUp(x, y).y + SourceY)
 Next y
Next x

'Show our Lens
SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicNew(0, 0), Binfo, DIB_RGB_COLORS
Pic.Refresh
Drawing = False
DoEvents
End Sub




Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
 If Drawing = False And Timer1.Enabled = False Then DrawLens x, Pic.Height - y
End Sub

'Show the Startpicture
Private Sub ShowPic()
 'copy the dib we got for our hidden picture to the front and clear the front
 SetDIBits Pic.hdc, Pic.Image.Handle, 0, Binfo.bmiHeader.biHeight, PicOrg(0, 0), Binfo, DIB_RGB_COLORS
 Pic.Refresh
End Sub


Private Sub CreateLens(ByRef LensArray() As PointApi, Diameter As Long, Magnification As Single)
 'Some simple Math :o)
 'to calculate our Lens

 'Thx to
 'Joey ????
 '   and
 'Jeff Lawson
 'for the Infos they posted about Lenscalculations

 Dim Radius As Integer
 Dim Sphere As Single
 Dim x As Long
 Dim y As Long
 Dim XOld As Long
 Dim YOld As Long
 Dim XNew As Long
 Dim YNew As Long
 Dim Z As Long
 Dim A As Long
 Dim B As Long
 Dim tmp1 As Long
 Dim tmp As Long

 Radius = Diameter / 2

 Sphere = Sqr(Radius ^ 2 - Magnification ^ 2)

 ReDim LensArray(Diameter, Diameter)


 For x = -Radius To -Radius + Diameter - 1
  For y = -Radius To -Radius + Diameter - 1
   If x ^ 2 + y ^ 2 >= Sphere ^ 2 Then
    A = x
    B = y
   Else
    Z = Sqr(Radius ^ 2 - x ^ 2 - y ^ 2)
    A = Int(x * Magnification / Z + 0.5)
    B = Int(y * Magnification / Z + 0.5)
   End If
   tmp1 = (1 + (y + Radius) * Diameter + (x + Radius))
   YOld = CInt(tmp1 / Diameter - 0.5)
   XOld = CInt(tmp1 - YOld * Diameter)
   tmp = (B + Radius) * Diameter + (A + Radius)
   YNew = CInt(tmp / Diameter - 0.5)
   XNew = CInt(tmp - YNew * Diameter)
   If XNew = 200 Then
    x = x
   End If
   LensArray(XOld, YOld).x = XNew
   LensArray(XOld, YOld).y = YNew
  Next y
 Next x
End Sub

'Move Our Lens automatic
Private Sub Timer1_Timer()
 Dim Speed As Byte
 'Scrollspeed
 Speed = 5

 If Drawing = False Then
  If DirX = 1 Then
   If MoveX < Pic.Width - HSSIZE / 2 Then
    MoveX = MoveX + Speed
   Else
    DirX = 0
    MoveX = MoveX - Speed
   End If
  Else
   If MoveX > HSSIZE / 2 Then
    MoveX = MoveX - Speed
   Else
    DirX = 1
    MoveX = MoveX + Speed
   End If
  End If
  If DirY = 1 Then
   If MoveY < Pic.Height - HSSIZE / 2 Then
    MoveY = MoveY + Speed
   Else
    DirY = 0
    MoveY = MoveY - Speed
   End If
  Else
   If MoveY > HSSIZE / 2 Then
    MoveY = MoveY - Speed
   Else
    DirY = 1
    MoveY = MoveY + Speed
   End If
  End If
  DrawLens MoveX, MoveY
 End If
End Sub

Private Sub Timer2_Timer()
Dim tmp As Boolean
 'if we change the Lenssize while the code draws then
 'dont calculate a new
 If Drawing = False Then
  tmp = Timer1.Enabled
  'Turn Createtimer off
  Timer2.Enabled = False
  'Turn Movetimer off
  Timer1.Enabled = False
  'Create new Lens
  CreateLens LookUp, HSSIZE, HSMAGN
  'Turn Movetimer on
  Timer1.Enabled = tmp
 End If
End Sub

'Test if we are in ide or compiled mode
'Private Function InIde() As Boolean
' On Error GoTo DivideError
' Debug.Print 1 / 0
' Exit Function
'DivideError:
' InIde = True
'End Function

'Private Sub ChkAutomove_Click()
' If ChkAutomove.Value = 0 Then Timer1.Enabled = False Else Timer1.Enabled = True
'End Sub

