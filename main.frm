VERSION 4.00
Begin VB.Form frmMain 
   Caption         =   "Demo Form"
   ClientHeight    =   5310
   ClientLeft      =   1425
   ClientTop       =   2565
   ClientWidth     =   5895
   Height          =   5715
   Icon            =   "main.frx":0000
   Left            =   1365
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   393
   Top             =   2220
   Width           =   6015
   Begin VB.Timer TimerStartup 
      Interval        =   100
      Left            =   5280
      Top             =   4680
   End
   Begin VB.CheckBox ChkRotateY 
      Caption         =   "Rotate Y"
      Height          =   375
      Left            =   480
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
   Begin MSComDlg.CommonDialog DlgOpenObj 
      Left            =   4560
      Top             =   4680
      _Version        =   65536
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   0
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

Dim gltree As GLObj
Private Sub GLDrawTree()

    Dim i As Integer
    Dim j As Integer
    Dim VertexIdx As Integer
    Dim MaterialIdx As Integer
    
    If 0 < gltree.FacesSz Then
    For i = 0 To gltree.FacesSz - 1
        If 3 = gltree.Faces(i).VertexIdxSz Then
        glBegin GL_TRIANGLES
        
        MaterialIdx = gltree.Faces(i).MaterialIdx
        glColor3f _
            gltree.Materials(MaterialIdx).Diffuse(0), _
            gltree.Materials(MaterialIdx).Diffuse(1), _
            gltree.Materials(MaterialIdx).Diffuse(2)
        
        For j = 0 To gltree.Faces(i).VertexIdxSz - 1
            VertexIdx = gltree.Faces(i).VertexIdx(j)
            If 0 < VertexIdx Then
                glNormal3f _
                    gltree.Vertices(VertexIdx).x, _
                    gltree.Vertices(VertexIdx).y, _
                    gltree.Vertices(VertexIdx).z
                glVertex3f _
                    gltree.Vertices(VertexIdx).x, _
                    gltree.Vertices(VertexIdx).y, _
                    gltree.Vertices(VertexIdx).z
            End If
        Next j
        glEnd
        End If
    Next i
    End If
End Sub

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
    
    RetVal = 1
    While 0 <> RetVal
        GLhDC = GetDC(PictureGL.hwnd)
    
        PFormatI = ChoosePixelFormat(GLhDC, PFormat)
        If 1 <> SetPixelFormat(GLhDC, PFormatI, PFormat) Then
            MsgBox "Error setting pixel format: " & GetLastError, vbCritical, "OpenGL Error"
            End
        End If
        
        frmLog.LogLine "hDC: " & GLhDC
        GLhRC = wglCreateContext(GLhDC)
        frmLog.LogLine "hRC: " & GLhRC
        If 0 = GLhRC Then
            'Problem setting up hRC, so skip rest of setup and start again.
            frmLog.LogLine "Error creating hRC: " & GetLastError
        Else
            wglMakeCurrent GLhDC, GLhRC
            'wglMakeCurrent just seems to always return 1 even if it failed?
            RetVal = GetLastError
            'If 0 <> RetVal Then
            '    MsgBox "Error creating context: " & RetVal, vbCritical, "OpenGL Error"
            '    End
            'End If
        End If
    Wend
    
    glViewport 0, 0, 320, 240

    With dlgopenobj
        .DialogTitle = "Open Model"
        .Filter = "Wavefront Object File (.obj)|*.obj"
        .ShowOpen
        If 0 = Len(.filename) Then End
        GLLoadObj gltree, .filename
    End With
    
    GLViewObjTree gltree
        
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

Private Sub Form_Unload(Cancel As Integer)
    TimerGL.Enabled = False
    wglMakeCurrent GLhDC, vbNull
    wglDeleteContext GLhRC
    End
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
    
    'GLCube
    GLDrawTree
    
    glPopMatrix
    
    glFlush
    SwapBuffers GLhDC
    'frmMain.Caption = "" & (Val(frmMain.Caption) + 1)
    
    If ChkRotateY.Value Then
        HScrollY.Value = HScrollY.Value + 5
        If HScrollY.Value >= 355 Then HScrollY.Value = 0
    End If
End Sub


Private Sub TimerStartup_Timer()
    GLStart
    TimerStartup.Enabled = False
End Sub
