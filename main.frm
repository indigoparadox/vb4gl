VERSION 4.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   4245
   ClientLeft      =   3375
   ClientTop       =   2070
   ClientWidth     =   5670
   Height          =   4650
   Left            =   3315
   LinkTopic       =   "Form1"
   ScaleHeight     =   283
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   378
   Top             =   1725
   Width           =   5790
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   3720
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   2880
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

Dim RotateX As Long
Dim RotateY As Long
Dim TranslateZ As Long

Private Sub Form_Load()
    
    Dim hrc As Long
    Dim PFormat As PIXELFORMATDESCRIPTOR
    Dim PFormatI As Long
    Dim Aspect As Single
    
    TranslateZ = -5
    
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

    PFormatI = ChoosePixelFormat(frmMain.hdc, PFormat)
    If 1 <> SetPixelFormat(frmMain.hdc, PFormatI, PFormat) Then
        MsgBox "Error setting pixel format: " & GetLastError, vbCritical, "OpenGL Error"
    Else
        hrc = wglCreateContext(frmMain.hdc)
        If 1 = wglMakeCurrent(frmMain.hdc, hrc) Then
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
        Else
            MsgBox "Error creating context: " & GetLastError, vbCritical, "OpenGL Error"
        End If
    End If
    
    'ShowGLError
    
End Sub



Private Sub Timer1_Timer()
    glClearColor 0#, 0#, 0#, 1#
    glClear 16384
    
    glPushMatrix
    glTranslatef 0#, 0#, TranslateZ
    glRotatef RotateX, 1#, 0#, 0#
    glRotatef RotateY, 0#, 1#, 0#
    
    GLCube
    glPopMatrix
    
    glFlush
    SwapBuffers (frmMain.hdc)
    'frmMain.Caption = "" & (Val(frmMain.Caption) + 1)
    
    RotateY = RotateY + 5
    
End Sub


