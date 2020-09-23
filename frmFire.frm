VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Fire"
   ClientHeight    =   3930
   ClientLeft      =   4185
   ClientTop       =   2850
   ClientWidth     =   4815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   262
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   321
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   3660
      Width           =   735
   End
   Begin VB.PictureBox picFire 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   0
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   0
      Width           =   4800
   End
   Begin VB.Label lblFPS 
      AutoSize        =   -1  'True
      Caption         =   "0 FPS"
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   3660
      Width           =   435
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DX As New DirectX7
Dim DD As DirectDraw7

'type used to determine the size of the picturebox
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

'Used to speed up the copying of the coolingmap
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'this will be used to get FPS
Private Declare Function GetTickCount Lib "kernel32" () As Long
'used to get the bitmap information from picturebox
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'sets the pixel colors in the picturebox
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
'used to hide the mouse
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'***************************
'To change fire size, you must change the
'picturebox size, fwidth, fheight, and the
'Buffer1 and CoolingMap to match them

'width of the fire area
Const fWidth = 320
'height of the fire area
Const fHeight = 240

'The number of bits to use for color
'16 is faster
'32 is slower, but looks better
Const ColorMode = 32
'This is Height * Width of the picture box
Const fHxW = 76800
'This is used for the cooling map
Const MemSize = fHxW - fWidth
'This is used for the cooling map also
Const fW1 = fWidth + 1
'holds the luminance of each pixel; 1 to fwidth * fheight
Dim Buffer1(1, 1 To fHxW) As Integer
'holds the cooling amount of each pixel; 1 to fwidth * fheight
Dim CoolingMap(1 To fHxW) As Integer

'holds the full color of flame
Dim FireColor(255) As Long

'used to get the color format (32 or 16 bit)
Dim PicInfo As BITMAP
'used by the program to determine 32/16 bit
Dim is32Bit As Boolean

'used to determine if the fire is running
Dim Running As Boolean
'used to determine if the user wants to stop fire
Dim StopIt As Boolean

'Holds which of the 2 fire buffers is holding current data
Dim CurBuf As Integer
'Holds which buffer will hold the new fire data
Dim NewBuf As Integer

'to control the fire intensity
Dim MinFire As Long
'to control cooling map smoothing
Dim MinCool As Long
'used in the loops
Dim I As Long
'the maximum loop count
Dim MaxInf As Long
'the minimum loop count
Dim MinInf As Long
'how many total pixels there are
Dim TotInf As Long

'Used to hold the 16-bit picture
Dim PicBitsI() As Integer
'Used to hold the 32-bit picture
Dim PicBitsL() As Long

Private Sub cmdStart_Click()
    
    'checks to see if the loop is already running
    If Running = True Then
        'if running, then stop it
        StopIt = True
        ShowCursor 1
    'if not running then lets start
    Else
        ShowCursor 0
        'let everything know it is running
        Running = True
        'we don't want to stopit, we just started it
        StopIt = False
        'change the command so user knows to click to stop
        cmdStart.Caption = "Stop"
        
        'Start the fireloop
        If is32Bit Then
            Call DoFire_32bit
        Else
            Call DoFire_16bit
        End If
        
        'loop is stopped so we don't need to stop it anymore
        StopIt = False
        'loop isn't running anymore
        Running = False
        'let user know to click command to start fire up
        cmdStart.Caption = "Start"
        'end the if statement from above (beginning of sub)
    
        'Restore the old display mode
        Call DD.RestoreDisplayMode
        Call DD.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    End If
End Sub

Private Sub cmdStart_KeyPress(KeyAscii As Integer)
    Form_KeyPress (KeyAscii)
End Sub

Private Sub Form_Activate()
    'Get the information about the picturebox
    GetObject picFire.Image, Len(PicInfo), PicInfo
    
    'Check the number of bits/pixel must be 16 or 32
    If PicInfo.bmBitsPixel = 16 Then
        '16-bit color
        is32Bit = False
        
        'Setup the array so it can hold a picture
        ReDim PicBitsI(1 To fHxW * (PicInfo.bmBitsPixel / 8)) As Integer
        ReDim PicBitsL(0 To 0) As Long
        
        'get what the maximum value for our fire loop needs to be
        MaxInf = (UBound(PicBitsI) / 2) - fWidth - 1
        
        'find out how many pixels there are in total
        TotInf = UBound(PicBitsI) / 2 - 1

    ElseIf PicInfo.bmBitsPixel = 32 Then
        '32-bit color
        is32Bit = True
        
        'Setup the array so it can hold a picture
        ReDim PicBitsL(1 To fHxW * (PicInfo.bmBitsPixel / 8)) As Long
        ReDim PicBitsI(0 To 0) As Integer
        'get what the maximum value for our fire loop needs to be
        MaxInf = (UBound(PicBitsL) / 4) - fWidth - 1
        'find out how many pixels there are in total
        TotInf = UBound(PicBitsL) / 4 - 1
    Else
        'You don't meet the requirements to run the program
        MsgBox "Sorry, your display settings are not currently supported!"
        End
    End If

    'the program isn't running
    Running = False
    
    'the loop isn't running so don't need to stop
    StopIt = False

    'get what the minimum value for our fire loop needs to be
    MinInf = fWidth + 1
    'sets the place to end cooling map control
    MinCool = MaxInf - (fWidth * 4)
    'sets the intensity of the fire
    MinFire = fWidth * 4
    
    'Setup the color array of the fire
    SetColorArrays
    
    'add some hotspots to start
    AddHotSpots (600)
    
    'add some coldspots to start
    AddColdSpots (250)
    
    cmdStart_Click
End Sub

Public Sub SetColorArrays()
'this function just sets the RGB colors used in
'the flame
    Dim mPath As String
    
    'Get the location of the exe (hopefully the path of the pals as well)
    mPath = App.Path
    If Right(mPath, 1) <> "\" Then mPath = mPath & "\"
    
    'Open and load the correct fire pal
    If is32Bit = True Then
        Open mPath & "fire32.pal" For Binary As #1
            Get #1, , FireColor
        Close #1
    Else
        Open mPath & "fire16.pal" For Binary As #1
            Get #1, , FireColor
        Close #1
    End If
    
End Sub

Public Sub AddColdSpots(ByVal Number As Long)
'adds cooling spots so the flame cools unevenly
'variable for the for loop
    'Dim I As Long
    'sets up the randomize function
    Randomize Timer
   
    'start the loop
    For I = 1 To Number
        'Buffer1(CurBuf, TotInf - Int(Rnd * fWidth) - fWidth) = Int(Rnd * 8) + 247 '255 'Int(Rnd * 191) + 64
        'creates a cool pixel placed randomly with a random cooling amount
        CoolingMap(TotInf - Int(Rnd * MinFire) - fWidth) = Int(Rnd * 10)
        'CoolingMap(Int(Rnd * TotInf) + 1) = Int(Rnd * 200) + 10
    'end of loop
    Next I
    'holds the cooling pixel to the right of current
    Dim N1 As Integer
    'holds the cooling pixel to the left of current
    Dim N2 As Integer
    'holds the cooling pixel down from the current
    Dim N3 As Integer
    'holds the cooling pixel up from the current
    Dim N4 As Integer
    'starts the loop (don't need edges)
    For I = MaxInf To MinCool Step -1
        'gets the pixels to the right value
        N1 = CoolingMap(I + 1)
        'gets the pixels to the left value
        N2 = CoolingMap(I - 1)
        'gets the pixels underneath value
        N3 = CoolingMap(I + fWidth)
        'gets the pixels above value
        N4 = CoolingMap(I - fWidth)
        'gets the average of the pixels around it
        CoolingMap(I) = CByte((N1 + N2 + N3 + N4) / 4)
    'end of loop
    Next I

    For I = 1 To TotInf - fWidth
        'copy the pixels back but up one pixel
        CoolingMap(I) = CoolingMap(I + fWidth)
    'end of loop
    Next I

End Sub

Public Sub AddHotSpots(ByVal Number As Long)
'add hot spots so the flame grows from the bottom
'for the loop
    'Dim I As Long
    'setup the randomize function
    Randomize Timer
    'start of the for loop
    For I = 1 To Number
        'adds a hotspot to the bottom with a random value
        '1 used to fwidth
        Buffer1(CurBuf, TotInf - Int(Rnd * MinFire) - fWidth) = Int(Rnd * 8) + 247 '255 'Int(Rnd * 191) + 64
    'end of loop
    Next I
End Sub

Public Sub DoFire_32bit()
'this sub is used if the user is in 32-bit color
'holds the starting time (for FPS)
    Dim St As Long
    'holds the ending time (for FPS)
    Dim Et As Long
    'holds the luminance of pixel to the right
    Dim N1 As Integer
    'holds the luminance of pixel to the left
    Dim N2 As Integer
    'holds the luminance of pixel underneath
    Dim N3 As Integer
    'holds the luminance of pixel above
    Dim N4 As Integer
    'holds a value used in use with the picture
    Dim Counter As Long
    'holds how many frames have been done
    Dim Frames As Long
    'holds the value of the current buffer (see later)
    Dim OldBuf As Byte
    'holds the new luminance of the pixel
    Dim P As Integer
    'holds the cooling value of the pixel
    Dim Col As Integer
    'used for loop
    'Dim I As Long
    'gets the current time
    St = GetTickCount
    'sets the frames to 0 cuz we just started
    Frames = 0

    'start the loop
    Do
        'set the counter to 1
        Counter = 1
        'start loop to calculate the fire
        For I = MinInf To MaxInf
            'gets the luminance of the pixel to the right
            N1 = Buffer1(CurBuf, I + 1)
            'gets the luminance of the pixel to the left
            N2 = Buffer1(CurBuf, I - 1)
            'gets the luminance of the pixel underneath
            N3 = Buffer1(CurBuf, I + fWidth)
            'gets the luminance of the pixel above
            N4 = Buffer1(CurBuf, I - fWidth)
            
            'gets the cooling amount
            Col = CoolingMap(I)
            'finds the average of surrounding pixels - cooling amount
            P = CByte((N1 + N2 + N3 + N4) / 4) - Col
            'if value is less than 0 make it 0
            If P < 0 Then P = 0
            
            'sets the new color into the buffer
            Buffer1(NewBuf, I - fWidth) = P
            
            'update the color in the picture
            PicBitsL(Counter) = FireColor(Buffer1(NewBuf, I - fWidth))
            'increment the counter
            Counter = Counter + 1
        'end of loop
        Next I
        'we need to swap the buffers
        'this holds the current newbuf value
        OldBuf = NewBuf
        'sets the newbuf to the curbuf value
        NewBuf = CurBuf
        'sets the curbuf to the newbuf value (held in OldBuf)
        CurBuf = OldBuf
        'adds some hotspots
        AddHotSpots (10)
        'adds some coldspots
        AddColdSpots (125)
        'draws the new image
        SetBitmapBits picFire.Image, UBound(PicBitsL), PicBitsL(1)
        'updates the picturebox
        picFire.Refresh
        'allows the loop to see changes in the StopIt variable
        DoEvents
        'adds one to frames
        Frames = Frames + 1
    'continue loop until StopIt doesn't equal false
    Loop While StopIt = False
    
    'gets the current time
    Et = GetTickCount()
    'calculates the frames per second and displays them
    lblFPS.Caption = Format(Frames / ((Et - St) / 1000), "0.00") & " FPS"
End Sub

Public Sub DoFire_16bit()
'this sub is used if the user is in 32-bit color
'holds the starting time (for FPS)
    Dim St As Long
    'holds the ending time (for FPS)
    Dim Et As Long
    'holds the luminance of pixel to the right
    Dim N1 As Long
    'holds the luminance of pixel to the left
    Dim N2 As Long
    'holds the luminance of pixel underneath
    Dim N3 As Long
    'holds the luminance of pixel above
    Dim N4 As Long
    'holds a value used in use with the picture
    Dim Counter As Long
    'holds how many frames have been done
    Dim Frames As Long
    'holds the value of the current buffer (see later)
    Dim OldBuf As Byte
    'holds the new luminance of the pixel
    Dim P As Integer
    'holds the cooling value of the pixel
    Dim Col As Integer
    'used for loops
    'Dim I As Long
    'gets the current time
    St = GetTickCount
    'sets the frames to 0 cuz we just started
    Frames = 0
    
    'start the loop
    Do
        'set the counter to 1
        Counter = 1
        'start loop to calculate the fire
        For I = MinInf To MaxInf
            'gets the luminance of the pixel to the right
            N1 = Buffer1(CurBuf, I + 1)
            'gets the luminance of the pixel to the left
            N2 = Buffer1(CurBuf, I - 1)
            'gets the luminance of the pixel underneath
            N3 = Buffer1(CurBuf, I + fWidth)
            'gets the luminance of the pixel above
            N4 = Buffer1(CurBuf, I - fWidth)
            
            'gets the cooling amount
            Col = CoolingMap(I)
            'finds the average of surrounding pixels - cooling amount
            P = CByte((N1 + N2 + N3 + N4) / 4) - Col
            'if value is less than 0 make it 0
            If P < 0 Then P = 0
            
            'sets the new color into the buffer
            Buffer1(NewBuf, I - fWidth) = P
            
            'update the color in the picture
            PicBitsI(Counter) = FireColor(Buffer1(NewBuf, I - fWidth))
            'increment the counter
            Counter = Counter + 1
        'end of loop
        Next I
        'we need to swap the buffers
        'this holds the current newbuf value
        OldBuf = NewBuf
        'sets the newbuf to the curbuf value
        NewBuf = CurBuf
        'sets the curbuf to the newbuf value (held in OldBuf)
        CurBuf = OldBuf
        'adds some hotspots
        AddHotSpots (10)
        'adds some coldspots
        AddColdSpots (125)
        'draws the new image
        SetBitmapBits picFire.Image, UBound(PicBitsI), PicBitsI(1)
        'updates the picturebox
        picFire.Refresh
        'allows the loop to see changes in the StopIt variable
        DoEvents
        'adds one to frames
        Frames = Frames + 1
    'continue loop until StopIt doesn't equal false
    Loop While StopIt = False
    
    'gets the current time
    Et = GetTickCount()
    'calculates the frames per second and displays them
    lblFPS.Caption = Format(Frames / ((Et - St) / 1000), "0.00") & " FPS"
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        If Running = True Then
            cmdStart_Click
        End If
        End
    End If
End Sub

Private Sub Form_Load()
    'Create the DirectDraw surface
    Set DD = DX.DirectDrawCreate("")
    Me.Show
    
    'change the display size
    Call DD.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
    DD.SetDisplayMode 320, 240, ColorMode, 0, DDSDM_DEFAULT
End Sub

Private Sub Form_Terminate()
    If Running = True Then
        cmdStart_Click
    End If
    End
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Running = True Then
        cmdStart_Click
    End If
    End
End Sub

Private Sub picFire_KeyPress(KeyAscii As Integer)
    Form_KeyPress (KeyAscii)
End Sub
