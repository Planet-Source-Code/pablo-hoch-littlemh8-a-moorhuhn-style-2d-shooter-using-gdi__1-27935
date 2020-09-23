VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "littleMH8 - my second gdi-based game ;)"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   353
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   595
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00404040&
      Height          =   3975
      Left            =   600
      MouseIcon       =   "frmGame.frx":030A
      MousePointer    =   2  'Kreuz
      ScaleHeight     =   261
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   509
      TabIndex        =   0
      Top             =   480
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   3975
      Left            =   600
      ScaleHeight     =   3915
      ScaleWidth      =   7635
      TabIndex        =   2
      Top             =   480
      Width           =   7695
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "if you want to play again, you have to restart this game."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1080
         TabIndex        =   9
         Top             =   3480
         Width           =   5415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "...and directx requires many dlls."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   2400
         Width           =   5175
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "of course directx is better, but gdi is easier ;-)"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   2160
         Width           =   5175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "this is my second gdi-based game."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1920
         Width           =   5175
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "written by mel aka pablo hoch"
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1920
         TabIndex        =   5
         Top             =   1680
         Width           =   5175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "a melaxis.com-game."
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   5175
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Zentriert
         BackColor       =   &H00000000&
         Caption         =   "littleMH8"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1695
         Left            =   120
         TabIndex        =   3
         Top             =   -120
         Width           =   7455
      End
   End
   Begin VB.Timer tmrTimer 
      Interval        =   1000
      Left            =   120
      Top             =   3600
   End
   Begin VB.Timer tmrBild 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   120
      Top             =   3000
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   120
      Top             =   2400
   End
   Begin VB.Timer tmrNewOpfer 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "Welcome to litleMH8, a moorhuhn-style 2d-shooter :) enjoy this simple game :DD"
      Height          =   255
      Left            =   600
      TabIndex        =   1
      Top             =   4560
      Width           =   7695
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

'Constants for the GenerateDC function
'**LoadImage Constants**
Const IMAGE_BITMAP As Long = 0
Const LR_LOADFROMFILE As Long = &H10
Const LR_CREATEDIBSECTION As Long = &H2000
'****************************************

'DCs
Private dcBackGround As Long
Private dcCursor As Long
Private dcCursorMask As Long
Private dcPatrone As Long
Private dcOpfer1 As Long
Private dcOpfer1Mask As Long
Private dcOpfer2 As Long
Private dcOpfer2Mask As Long
'Status...
Private CursorVisible As Boolean
Private MousePos As POINTAPI
Private Patronen As Byte
Private Score As String
Private TimeLeft As String

Private Type Opfer
    X As Long
    Y As Long
    pic As Integer
    Alive As Boolean
    Visible As Boolean
End Type

Private Opfers() As Opfer


'IN: FileName: The file name of the graphics
'OUT: The Generated DC
Public Function GenerateDC(FileName As String) As Long
Dim DC As Long
Dim hBitmap As Long

'Create a Device Context, compatible with the screen
DC = CreateCompatibleDC(0)

If DC = 0 Then    'DC < 1 Then
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 1
    Exit Function
End If

'Load the image....BIG NOTE: This function is not supported under NT, there you can not
'specify the LR_LOADFROMFILE flag
hBitmap = LoadImage(0, FileName, IMAGE_BITMAP, 0, 0, LR_LOADFROMFILE Or LR_CREATEDIBSECTION)

If hBitmap = 0 Then 'Failure in loading bitmap
    DeleteDC DC
    GenerateDC = 0
    'Raise error
    Err.Raise vbObjectError + 2
    Exit Function
End If

'Throw the Bitmap into the Device Context
SelectObject DC, hBitmap

'Return the device context
GenerateDC = DC

'Delte the bitmap handle object
DeleteObject hBitmap

End Function
'Deletes a generated DC
Private Function DeleteGeneratedDC(DC As Long) As Long

If DC > 0 Then
    DeleteGeneratedDC = DeleteDC(DC)
Else
    DeleteGeneratedDC = 0
End If

End Function


Private Sub Form_Load()
ReDim Opfers(0)
'pics laden
dcBackGround = GenerateDC(App.Path & "\data\gfx\bg.bmp")
dcCursor = GenerateDC(App.Path & "\data\gfx\cursor.bmp")
dcCursorMask = GenerateDC(App.Path & "\data\gfx\cursor_mask.bmp")
dcPatrone = GenerateDC(App.Path & "\data\gfx\patrone.bmp")
dcOpfer1 = GenerateDC(App.Path & "\data\gfx\opfer1.bmp")
dcOpfer1Mask = GenerateDC(App.Path & "\data\gfx\opfer1_mask.bmp")
dcOpfer2 = GenerateDC(App.Path & "\data\gfx\opfer2.bmp")
dcOpfer2Mask = GenerateDC(App.Path & "\data\gfx\opfer2_mask.bmp")
'einstellungen
Patronen = 5
Score = 0
TimeLeft = 90
'background zeichnen ;-)
BitBlt pic.hdc, 0, 75, pic.ScaleWidth, pic.ScaleHeight, dcBackGround, 0, 0, vbSrcCopy
'patronen zeichnen ;)
For i = 0 To (Patronen - 1)
    BitBlt pic.hdc, 5 + (i * 18), 215, 16, 43, dcPatrone, 0, 0, vbSrcCopy
Next
tmrGame.Enabled = True
tmrNewOpfer.Enabled = True
tmrBild.Enabled = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CursorVisible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
DeleteGeneratedDC dcBackGround
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CursorVisible = False
End Sub

Private Sub pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If Button = vbLeftButton Then
    'ballern ;)
    If Patronen > 0 Then
        'treffer?
        For i = LBound(Opfers) To UBound(Opfers)
            'nur sichtbare!
            If Opfers(i).Visible Then
                'und nur lebende :D
                If Opfers(i).Alive Then
                    'checken ob getroffen
                    If (X > Opfers(i).X) And (X < (Opfers(i).X + 71)) Then
                        'okay x stimmt schonmal.... und was is mit y?
                        If (Y > Opfers(i).Y) And (Y < (Opfers(i).Y + 43)) Then
                            ' auch das! also ein treffer :-)
                            Opfers(i).Alive = False
                            Score = Score + 25
                        End If
                    End If
                End If
            End If
        Next
        Patronen = Patronen - 1
    Else
        'keine muni mehr
    End If
ElseIf Button = vbRightButton Then
    'nachladen
    Patronen = 5
End If
Draw
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
CursorVisible = True
MousePos.X = X
MousePos.Y = Y
Draw
End Sub


Sub Draw()
On Error Resume Next                'das entladen des letzten objekts könnte problem bereiten ;)
Dim val As Integer
Randomize
val = Rnd * 1
'hintergrund kommt immer zu erst :D
pic.Cls
BitBlt pic.hdc, 0, 75, pic.ScaleWidth, pic.ScaleHeight, dcBackGround, 0, 0, vbSrcCopy
'''''...und zum schluss die maus
''''BitBlt pic.hdc, MousePos.X, MousePos.Y, 43, 43, dcCursorMask, 0, 0, vbSrcAnd
''''BitBlt pic.hdc, MousePos.X, MousePos.Y, 43, 43, dcCursor, 0, 0, vbSrcPaint
'''''geht alles net ;((
'die opfer kommen jetzt dran...
For i = LBound(Opfers) To UBound(Opfers)
    'nur noch aktive opfer malen (d.h. die aufm screen sind)
    If Opfers(i).Visible Then
        If Opfers(i).Alive Then
            'fliegend
            Opfers(i).X = Opfers(i).X + 1
            If val = 1 Then
                Randomize
                val = Rnd * 1
                If val = 1 Then
                    Opfers(i).Y = Opfers(i).Y - 1
                Else
                    Opfers(i).Y = Opfers(i).Y + 1
                End If
            End If
            If Opfers(i).X > (Me.ScaleWidth + 10) Then
                ' das war das letzte mal hehe
                Opfers(i).Visible = False
            End If
        Else
            'fallend
            Opfers(i).Y = Opfers(i).Y + 2
            If Opfers(i).Y > (Me.ScaleHeight + 10) Then
                ' in zukunft nicht mehr zeichnen :D
                Opfers(i).Visible = False
            End If
        End If
        'zeichnen
        If Opfers(i).pic = 1 Then
            BitBlt pic.hdc, Opfers(i).X, Opfers(i).Y, 71, 43, dcOpfer1Mask, 0, 0, vbSrcAnd
            BitBlt pic.hdc, Opfers(i).X, Opfers(i).Y, 71, 43, dcOpfer1, 0, 0, vbSrcPaint
        Else
            BitBlt pic.hdc, Opfers(i).X, Opfers(i).Y, 71, 43, dcOpfer2Mask, 0, 0, vbSrcAnd
            BitBlt pic.hdc, Opfers(i).X, Opfers(i).Y, 71, 43, dcOpfer2, 0, 0, vbSrcPaint
        End If
        'nächstes mal nicht mehr?
        If Opfers(i).Visible = False Then
            'letztes flugobjekt?
            If i = UBound(Opfers) Then
                'jo, entladen
                ReDim Preserve Opfers(0 To (UBound(Opfers) - 1))
            End If
        End If
    End If
Next
'schließlich noch die patronen
If Patronen > 0 Then
    For i = 0 To (Patronen - 1)
        BitBlt pic.hdc, 5 + (i * 18), 215, 16, 43, dcPatrone, 0, 0, vbSrcCopy
    Next
End If
'...und die punktzahl :-)
TextOut pic.hdc, 460, 5, String(5 - Len(Score), "0") & Score, 5
TextOut pic.hdc, 7, 5, String(2 - Len(TimeLeft), "0") & TimeLeft, 2
'und dann anzeigen
pic.Refresh
If TimeLeft = 0 Then
    tmrNewOpfer.Enabled = False
    tmrGame.Enabled = False
    tmrBild.Enabled = False
    tmrTimer.Enabled = False
    Pause 500
    pic.Visible = False
End If
End Sub

Private Sub tmrBild_Timer()
On Error Resume Next
Dim val As Integer
'ändert das aussehen der opfer ;)
For i = LBound(Opfers) To UBound(Opfers)
    'natürlich nur sichtbare
    If Opfers(i).Visible Then
        If Opfers(i).Alive Then
            'tote machen nix mehr ;))
            Randomize
            val = (Rnd * 1) + 1
            Opfers(i).pic = val
        End If
    End If
Next
End Sub

Private Sub tmrGame_Timer()
Draw
End Sub

Private Sub tmrNewOpfer_Timer()
Dim val As Integer, id As Integer
Randomize
val = Rnd * 1
If val = 1 Then
    'neues opfer erstellen
    id = UBound(Opfers) + 1
    ReDim Preserve Opfers(0 To id)
    Opfers(id).Alive = True
    Opfers(id).X = -72
    Randomize
    Opfers(id).Y = Int(Rnd * (Me.ScaleHeight - 30))
    Opfers(id).pic = 1
    Opfers(id).Visible = True
    'gezeichnet wird das nächste mal, ist ja eh noch unsichtbar :D
End If
End Sub

Private Sub tmrTimer_Timer()
TimeLeft = TimeLeft - 1
End Sub




Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    
    Do
      u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub

Public Function ShortPath(Path As String) As String
On Error Resume Next
Dim kurz As String
kurz = Space$(265)
Call GetShortPathName(Path, kurz, Len(kurz))
kurz = Replace(kurz, Chr(0), "")
kurz = Trim(kurz)
ShortPath = kurz
End Function
