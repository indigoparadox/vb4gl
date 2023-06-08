VERSION 4.00
Begin VB.Form frmMain 
   Caption         =   "Demo Form"
   ClientHeight    =   5310
   ClientLeft      =   1755
   ClientTop       =   2160
   ClientWidth     =   5790
   Height          =   5715
   Icon            =   "main.frx":0000
   Left            =   1695
   LinkTopic       =   "Form1"
   ScaleHeight     =   354
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   386
   Top             =   1815
   Width           =   5910
   Begin VB.CheckBox ChkLight 
      Caption         =   "&Lighting"
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.VScrollBar VScrollY 
      Height          =   3615
      Left            =   120
      Max             =   20
      Min             =   -20
      TabIndex        =   4
      Top             =   360
      Width           =   255
   End
   Begin VB.Timer TimerStartup 
      Interval        =   100
      Left            =   5280
      Top             =   4680
   End
   Begin VB.CheckBox ChkRotateY 
      Caption         =   "&Rotate Y"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   4800
      Value           =   1  'Checked
      Width           =   1215
   End
   Begin VB.VScrollBar VScrollZ 
      Height          =   3615
      Left            =   5400
      Max             =   20
      Min             =   -20
      TabIndex        =   2
      Top             =   360
      Value           =   -5
      Width           =   255
   End
   Begin VB.HScrollBar HScrollYR 
      Height          =   255
      Left            =   480
      Max             =   360
      TabIndex        =   1
      Top             =   4080
      Width           =   4815
   End
   Begin VB.PictureBox PictureGL 
      Height          =   3600
      Left            =   480
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
    PFormat.cColorBits = 16
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
    PFormat.cDepthBits = 8
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
        frmLog.LogLine "hDC: " & GLhDC
    
        PFormatI = ChoosePixelFormat(GLhDC, PFormat)
        RetVal = GetLastError
        If 0 = PFormatI Or 0 <> RetVal Then
            frmLog.LogLine "Pixel Format: " & PFormatI
            GLShowSystemError "Error choosing pixel format", RetVal, False
            End
        Else
            frmLog.LogLine "Pixel Format: " & PFormatI
        End If
        
        SetPixelFormat GLhDC, PFormatI, PFormat
        RetVal = GetLastError
        If 0 <> RetVal Then
            GLShowSystemError "Error setting pixel format", RetVal, False
            End
        End If
        
        GLhRC = wglCreateContext(GLhDC)
        RetVal = GetLastError
        frmLog.LogLine "hRC: " & GLhRC
        If 0 = GLhRC Then
            'Problem setting up hRC, so skip rest of setup and start again.
            GLShowSystemError "Error creating hRC", RetVal, True
        Else
            wglMakeCurrent GLhDC, GLhRC
        End If
        
        DoEvents
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
    
    glEnable GL_LIGHTING
    
    glEnable GL_DEPTH_TEST
    glDepthMask GL_TRUE
    glDepthFunc GL_LESS
    glDepthRange 0, 1
    
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

    Dim LightPos(4) As Single
    Dim LightDir(3) As Single

    glClearColor 0#, 0#, 0#, 1#
    glClear GL_COLOR_BUFFER_BIT + GL_DEPTH_BUFFER_BIT
    
    If ChkLight.Value Then
        glEnable GL_LIGHTING
        If Not frmLights.Visible Then frmLights.Visible = True
        
        If 1 = frmLights.chklight0.Value Then
            glEnable GL_LIGHT0
            glEnable GL_COLOR_MATERIAL
            LightPos(0) = frmLights.slider0x.Value
            frmLights.lbl0px.Caption = Val(LightPos(0))
            LightPos(1) = frmLights.slider0y.Value
            frmLights.lbl0py.Caption = Val(LightPos(1))
            LightPos(2) = frmLights.slider0z.Value
            frmLights.lbl0pz.Caption = Val(LightPos(2))
            LightPos(3) = frmLights.slider0w.Value
            frmLights.lbl0pw.Caption = Val(LightPos(3))
            glLightfv GL_LIGHT0, GL_POSITION, LightPos
            
            'LightDir(0) = frmLights.slider0dx.Value
            'LightDir(1) = frmLights.slider0dy.Value
            'LightDir(2) = frmLights.slider0dz.Value
            'glLightfv GL_LIGHT0, GL_SPOT_DIRECTION, LightDir
        Else
            glDisable GL_LIGHT0
            glDisable GL_COLOR_MATERIAL
        End If
    Else
        glDisable GL_LIGHTING
        If frmLights.Visible Then frmLights.Visible = False
    End If
    
    glPushMatrix
    glTranslatef 0#, VScrollY.Value, VScrollZ.Value
    glRotatef RotateX, 1#, 0#, 0#
    glRotatef HScrollYR.Value, 0#, 1#, 0#
    
    'GLCube
    GLDrawTree
    
    glPopMatrix
    
    glFlush
    SwapBuffers GLhDC
    'frmMain.Caption = "" & (Val(frmMain.Caption) + 1)
    
    If ChkRotateY.Value Then
        HScrollYR.Value = HScrollYR.Value + 5
        If HScrollYR.Value >= 355 Then HScrollYR.Value = 0
    End If
End Sub


Private Sub TimerStartup_Timer()
    GLStart
    TimerStartup.Enabled = False
End Sub

