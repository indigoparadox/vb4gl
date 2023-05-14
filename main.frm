VERSION 4.00
Begin VB.Form frmMain 
   Caption         =   "Demo Form"
   ClientHeight    =   5310
   ClientLeft      =   1740
   ClientTop       =   2100
   ClientWidth     =   5895
   Height          =   5715
   Icon            =   "main.frx":0000
   Left            =   1680
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   Top             =   1755
   Width           =   6015
   Begin VB.CommandButton CmdStart 
      Caption         =   "&Start"
      Height          =   495
      Left            =   4320
      TabIndex        =   4
      Top             =   4680
      Width           =   1455
   End
   Begin VB.CheckBox ChkRotateY 
      Caption         =   "Rotate Y"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Value           =   1  'Checked
      Width           =   2175
   End
   Begin VB.VScrollBar VScrollZ 
      Height          =   3615
      Left            =   5280
      Max             =   20
      Min             =   -20
      TabIndex        =   2
      Top             =   360
      Value           =   -5
      Width           =   255
   End
   Begin VB.HScrollBar HScrollY 
      Height          =   255
      Left            =   360
      Max             =   360
      TabIndex        =   1
      Top             =   4200
      Width           =   4815
   End
   Begin VB.PictureBox PictureGL 
      Height          =   3600
      Left            =   360
      ScaleHeight     =   236
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   316
      TabIndex        =   0
      Top             =   360
      Width           =   4800
   End
   Begin VB.Timer TimerGL 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   5280
      Top             =   4080
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
Option Explicit

Dim RotateX As Long
Dim RotateY As Long

Dim GLhDC As Long
Dim GLhRC As Long

Dim GLStarted As Boolean
Private Sub GLStart()
    Dim PFormat As PIXELFORMATDESCRIPTOR
    Dim PFormatI As Long
    Dim Aspect As Single
    Dim RetVal As Long
        
    PFormat.nSize = 40
    PFormat.nVersion = 1
    PFormat.dwFlags = PFD_DRAW_TO_WINDOW + PFD_SUPPORT_OPENGL + PFD_DOUBLEBUFFER
    PFormat.iPixelType = PFD_TYPE_RGBA
    PFormat.cColorBits = 24
    PFormat.cRedBits = 0
    PFormat.cRedShift = 0
    PFormat.cGreenBits = 0
    PFormat.cGreenShift = 0
    PFormat.cBlueBits = 0
    PFormat.cBlueShift = 0
    PFormat.cAlphaBits = 0
    PFormat.cAlphaShift = 0
    PFormat.cAccumBits = 0
    PFormat.cAccumRedBits = 0
    PFormat.cAccumGreenBits = 0
    PFormat.cAccumBlueBits = 0
    PFormat.cAccumAlphaBits = 0
    PFormat.cDepthBits = 32
    PFormat.cStencilBits = 0
    PFormat.cAuxBuffers = 0
    PFormat.iLayerType = PFD_MAIN_PLANE
    PFormat.bReserved = 0
    PFormat.dwLayerMask = 0
    PFormat.dwVisibleMask = 0
    PFormat.dwDamageMask = 0
    
    GLhDC = PictureGL.hdc

    PFormatI = ChoosePixelFormat(GLhDC, PFormat)
    If 1 <> SetPixelFormat(GLhDC, PFormatI, PFormat) Then
        MsgBox "Error setting pixel format: " & GetLastError, vbCritical, "OpenGL Error"
        End
    End If
    
    'While 0 = GLhRC
    GLhRC = wglCreateContext(GLhDC)
    'MsgBox "" & GLhRC
    'Wend
    wglMakeCurrent GLhDC, GLhRC
    'wglMakeCurrent just seems to always return 1 even if it failed?
    RetVal = GetLastError
    If 0 <> RetVal Then
        MsgBox "Error creating context: " & RetVal, vbCritical, "OpenGL Error"
        End
    End If
    
    glViewport 0, 0, 320, 240

    glEnable GL_NORMALIZE
    glEnable GL_CULL_FACE
    'glEnable GL_DEPTH_TEST
    
    'glDepthMask GL_TRUE
    'glDepthFunc GL_LESS
    'glDepthRange 0, 1
    
    'Setup 3D projection.
    glMatrixMode GL_PROJECTION
    glLoadIdentity
    Aspect = 320 / 240
    glFrustum -0.5 * Aspect, 0.5 * Aspect, -0.5, 0.5, 0.5, 10
    glMatrixMode GL_MODELVIEW
    
    TimerGL.Enabled = True
    
    'ShowGLError
End Sub


Private Sub CmdStart_Click()
    GLStart
End Sub

Private Sub Form_Load()
    GLStart
End Sub

Private Sub HScrollY_GotFocus()
    ChkRotateY.Value = 0
End Sub

Private Sub TimerGL_Timer()
    glClearColor 0#, 0#, 0#, 1#
    glClear 16384
    
    glPushMatrix
    glTranslatef 0#, 0#, VScrollZ.Value
    glRotatef RotateX, 1#, 0#, 0#
    glRotatef HScrollY.Value, 0#, 1#, 0#
    
    GLCube
    glPopMatrix
    
    glFlush
    SwapBuffers GLhDC
    'frmMain.Caption = "" & (Val(frmMain.Caption) + 1)
    
    If ChkRotateY.Value Then
        HScrollY.Value = HScrollY.Value + 5
        If HScrollY.Value >= 355 Then HScrollY.Value = 0
    End If
End Sub


