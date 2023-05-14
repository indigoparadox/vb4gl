Attribute VB_Name = "basGL"
Option Explicit

Public Type PIXELFORMATDESCRIPTOR
    nSize As Integer
    nVersion As Integer
    dwFlags As String * 1
    iPixelType As String * 1
    cColorBits As String * 1
    cRedBits As String * 1
    cRedShift As String * 1
    cGreenBits As String * 1
    cGreenShift As String * 1
    cBlueBits As String * 1
    cBlueShift As String * 1
    cAlphaBits As String * 1
    cAlphaShift As String * 1
    cAccumBits As String * 1
    cAccumRedBits As String * 1
    cAccumGreenBits As String * 1
    cAccumBlueBits As String * 1
    cAccumAlphaBits As String * 1
    cDepthBits As String * 1
    cStencilBits As String * 1
    cAuxBuffers As String * 1
    iLayerType As String * 1
    bReserved As String * 1
    dwLayerMask As Long
    dwVisibleMask As Long
    dwDamageMask As Long
End Type

Public Const PFD_TYPE_RGBA = 0

Public Const PFD_MAIN_PLANE = 0
Public Const PFD_OVERLAY_PLANE = 1
Public Const PFD_UNDERLAY_PLANE = (-1)

Public Const PFD_DOUBLEBUFFER = &H1
Public Const PFD_STEREO = &H2
Public Const PFD_DRAW_TO_WINDOW = &H4
Public Const PFD_DRAW_TO_BITMAP = &H8
Public Const PFD_SUPPORT_GDI = &H10
Public Const PFD_SUPPORT_OPENGL = &H20
Public Const PFD_GENERIC_FORMAT = &H40
Public Const PFD_NEED_PALETTE = &H80
Public Const PFD_NEED_SYSTEM_PALETTE = &H100
Public Const PFD_SWAP_EXCHANGE = &H200
Public Const PFD_SWAP_COPY = &H400
Public Const PFD_SWAP_LAYER_BUFFERS = &H800
Public Const PFD_GENERIC_ACCELERATED = &H1000
Public Const PFD_SUPPORT_DIRECTDRAW = &H2000
Public Const PFD_DIRECT3D_ACCELERATED = &H4000
Public Const PFD_SUPPORT_COMPOSITION = &H8000
Public Const PFD_DEPTH_DONTCARE = &H20000000
Public Const PFD_DOUBLEBUFFER_DONTCARE = &H40000000
Public Const PFD_STEREO_DONTCARE = &H80000000

Public Const GL_PROJECTION = &H1701
Public Const GL_MODELVIEW = &H1700

Public Const GL_POINTS = &H0
Public Const GL_LINES = &H1
Public Const GL_LINE_LOOP = &H2
Public Const GL_LINE_STRIP = &H3
Public Const GL_TRIANGLES = &H4
Public Const GL_TRIANGLE_STRIP = &H5
Public Const GL_TRIANGLE_FAN = &H6
Public Const GL_QUADS = &H7
Public Const GL_QUAD_STRIP = &H8
Public Const GL_POLYGON = &H9

Public Const GL_NORMALIZE = &HBA1
Public Const GL_DEPTH_TEST = &HB71
Public Const GL_CULL_FACE = &HB44

Public Const GL_TRUE = 1
Public Const GL_FALSE = 0

Public Const GL_NEVER = &H200
Public Const GL_LESS = &H201
Public Const GL_EQUAL = &H202
Public Const GL_LEQUAL = &H203
Public Const GL_GREATER = &H204
Public Const GL_NOTEQUAL = &H205
Public Const GL_GEQUAL = &H206
Public Const GL_ALWAYS = &H207

Public Declare Function ChoosePixelFormat Lib "Gdi32" (ByVal hdc As Long, ppfd As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SetPixelFormat Lib "Gdi32" (ByVal hdc As Long, ByVal format As Long, ppfd As PIXELFORMATDESCRIPTOR) As Boolean
Public Declare Function SwapBuffers Lib "Gdi32" (ByVal hdc As Long) As Boolean

Public Declare Function wglCreateContext Lib "Opengl32" (ByVal hdc As Long) As Long
Public Declare Function wglMakeCurrent Lib "Opengl32" (ByVal hdc As Long, ByVal hrc As Long) As Boolean

Public Declare Sub glClearColor Lib "Opengl32" (ByVal r As Single, ByVal g As Single, ByVal b As Single, ByVal a As Single)
Public Declare Sub glClear Lib "Opengl32" (ByVal b As Long)
Public Declare Sub glFlush Lib "Opengl32" ()
Public Declare Function glGetError Lib "Opengl32" () As Long
Public Declare Sub glViewport Lib "Opengl32" (ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
Public Declare Sub glFrustum Lib "Opengl32" (ByVal left As Double, ByVal right As Double, ByVal bottom As Double, ByVal top As Double, ByVal zNear As Double, ByVal zFar As Double)
Public Declare Sub glLoadIdentity Lib "Opengl32" ()
Public Declare Sub glMatrixMode Lib "Opengl32" (ByVal mode As Long)
Public Declare Sub glTranslatef Lib "Opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glRotatef Lib "Opengl32" (ByVal angle As Single, ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glVertex3f Lib "Opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glVertex3i Lib "Opengl32" (ByVal x As Long, ByVal y As Long, ByVal z As Long)
Public Declare Sub glColor3f Lib "Opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glNormal3f Lib "Opengl32" (ByVal x As Single, ByVal y As Single, ByVal z As Single)
Public Declare Sub glBegin Lib "Opengl32" (ByVal poly As Long)
Public Declare Sub glEnd Lib "Opengl32" ()
Public Declare Sub glPushMatrix Lib "Opengl32" ()
Public Declare Sub glPopMatrix Lib "Opengl32" ()
Public Declare Sub glEnable Lib "Opengl32" (ByVal cap As Long)
Public Declare Sub glDepthMask Lib "Opengl32" (ByVal tf As Boolean)
Public Declare Sub glDepthFunc Lib "Opengl32" (ByVal tf As Boolean)
Public Declare Sub glDepthRange Lib "Opengl32" (ByVal near As Double, ByVal far As Double)

Public Declare Function GetLastError Lib "Kernel32" () As Long
Public Sub GLCube()

'BACK
glBegin GL_TRIANGLES
glNormal3f 0, 0, 1#
glColor3f 1#, 1#, 1#
glVertex3f 1#, -1#, 1#
glVertex3f 1#, 1#, 1#
glVertex3f -1#, 1#, 1#

glVertex3f -1#, 1#, 1#
glVertex3f -1#, -1#, 1#
glVertex3f 1#, -1#, 1#
glEnd

'RIGHT
glBegin GL_TRIANGLES
glNormal3f 1#, 0, 0
glColor3f 0, 1#, 1#
glVertex3f 1#, -1#, -1#
glVertex3f 1#, 1#, -1#
glVertex3f 1#, 1#, 1#

glVertex3f 1#, 1#, 1#
glVertex3f 1#, -1#, 1#
glVertex3f 1#, -1#, -1#
glEnd

'LEFT
glBegin GL_TRIANGLES
glNormal3f -1#, 0, 0
glColor3f 1#, 1#, 0
glVertex3f -1#, -1#, 1#
glVertex3f -1#, 1#, 1#
glVertex3f -1#, 1#, -1#

glVertex3f -1#, 1#, -1#
glVertex3f -1#, -1#, -1#
glVertex3f -1#, -1#, 1#
glEnd

'FRONT
glBegin GL_TRIANGLES
glNormal3f 0, 0, -1#
glColor3f 0, 0, 1#
glVertex3f -1#, -1#, -1#
glVertex3f -1#, 1#, -1#
glVertex3f 1#, 1#, -1#

glVertex3f 1#, 1#, -1#
glVertex3f 1#, -1#, -1#
glVertex3f -1#, -1#, -1#
glEnd

'TOP
glBegin GL_TRIANGLES
glNormal3f 0, 1#, 0
glColor3f 0, 1#, 0
glVertex3f 1#, 1#, 1#
glVertex3f 1#, 1#, -1#
glVertex3f -1#, 1#, -1#

glVertex3f -1#, 1#, -1#
glVertex3f -1#, 1#, 1#
glVertex3f 1#, 1#, 1#
glEnd

'BOTTOM
glBegin GL_TRIANGLES
glNormal3f 0, -1#, 0
glColor3f 1#, 0, 0
glVertex3f 1#, -1#, -1#
glVertex3f 1#, -1#, 1#
glVertex3f -1#, -1#, 1#

glVertex3f -1#, -1#, 1#
glVertex3f -1#, -1#, -1#
glVertex3f 1#, -1#, -1#
glEnd

End Sub


Public Function GLShowError() As Boolean

    Dim ErrVal As Long
    Dim ErrMsg As String
    
    ErrVal = glGetError

    Select Case ErrVal
    
    Case 1280
        ErrMsg = "Invalid enumerated argument."
        
    Case 1281
        ErrMsg = "Invalid value."
        
    Case 1282
        ErrMsg = "Invalid operation."
        
    Case 1283
        ErrMsg = "Stack overflow."
    
    Case 1284
        ErrMsg = "Stack underflow."

    Case 1285
        ErrMsg = "Out of memory."
        
    End Select
    
    If 0 <> ErrVal Then
        MsgBox ErrMsg, vbCritical, "OpenGL Error"
        ShowGLError = True
    End If
End Function





