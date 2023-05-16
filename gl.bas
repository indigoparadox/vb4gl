Attribute VB_Name = "basGL"
Option Explicit

Public Type GLVertex
    x As Single
    y As Single
    z As Single
End Type

Public Type GLVTexture
    u As Single
    V As Single
    w As Single
End Type

Public Type GLMaterial
    Ambient(4) As Single
    Diffuse(4) As Single
    Specular(4) As Single
    Emissive(4) As Single
    SpecularExp As Single
    Name As String
End Type

Public Type GLFace
    VertexIdx() As Integer
    VertexIdxSz As Integer
    VNormalIdx() As Integer
    VNormalIdxSz As Integer
    VTextureIdx() As Integer
    VTextureIdxSz As Integer
End Type

Public Type GLObj
    Vertices() As GLVertex
    VerticesSz As Integer
    VNormals() As GLVertex
    VNormalsSz As Integer
    VTextures() As GLVTexture
    VTexturesSz As Integer
    Faces() As GLFace
    FacesSz As Integer
    Materials() As GLMaterial
    MaterialsSz As Integer
End Type

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
Public Declare Function GetDC Lib "User32" (ByVal hwnd As Long) As Long

Public Declare Function wglCreateContext Lib "Opengl32" (ByVal hdc As Long) As Long
Public Declare Function wglMakeCurrent Lib "Opengl32" (ByVal hdc As Long, ByVal hrc As Long) As Boolean
Public Declare Function wglDeleteContext Lib "Opengl32" (ByVal hrc As Long) As Boolean

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

Public Sub GLViewObjTree(ByRef ObjIn As GLObj)
    Dim FacesNode As Node
    Dim FaceIdx As Integer
    Dim FaceVertexIdx As Integer
    Dim FaceNodeIter As Node
    Dim VertexNodeIter As Node
    Dim ObjVertexIdx As Integer
    Dim ParentNodeIter As Node
    
    Set FacesNode = frmObjTree.treeviewobj.Nodes.Add(, , , "Faces")
    FacesNode.Image = 1
    For FaceIdx = 0 To ObjIn.FacesSz - 1
        Set FaceNodeIter = frmObjTree.treeviewobj.Nodes.Add(1, tvwChild, , "Face " & FaceIdx)
        
        'Add nodes for face vertices.
        Set ParentNodeIter = frmObjTree.treeviewobj.Nodes.Add(FaceNodeIter, tvwChild, , "Vertices")
        For FaceVertexIdx = 0 To ObjIn.Faces(FaceIdx).VertexIdxSz - 1
            ObjVertexIdx = ObjIn.Faces(FaceIdx).VertexIdx(FaceVertexIdx)
            Set VertexNodeIter = frmObjTree.treeviewobj.Nodes.Add( _
                ParentNodeIter.Index, tvwChild, , "Vertex " & ObjVertexIdx)
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "X: " & _
                    ObjIn.Vertices(ObjVertexIdx).x
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "Y: " & _
                    ObjIn.Vertices(ObjVertexIdx).y
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "Z: " & _
                    ObjIn.Vertices(ObjVertexIdx).z
        Next FaceVertexIdx
        
        'Add nodes for face normals.
        Set ParentNodeIter = frmObjTree.treeviewobj.Nodes.Add(FaceNodeIter, tvwChild, , "Normals")
        For FaceVertexIdx = 0 To ObjIn.Faces(FaceIdx).VNormalIdxSz - 1
            ObjVertexIdx = ObjIn.Faces(FaceIdx).VNormalIdx(FaceVertexIdx)
            Set VertexNodeIter = frmObjTree.treeviewobj.Nodes.Add( _
                ParentNodeIter.Index, tvwChild, , "Normal " & ObjVertexIdx)
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "X: " & _
                    ObjIn.VNormals(ObjVertexIdx).x
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "Y: " & _
                    ObjIn.VNormals(ObjVertexIdx).y
            frmObjTree.treeviewobj.Nodes.Add _
                VertexNodeIter.Index, tvwChild, , "Z: " & _
                    ObjIn.VNormals(ObjVertexIdx).z
        Next FaceVertexIdx
    Next FaceIdx
    
    frmObjTree.Show
End Sub

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


Public Function GLLoadObj(ByRef ObjIn As GLObj, ByVal ObjPath As String) As Boolean

    Dim ObjFileNo As Integer
    Dim Line As String
    Dim LineArr() As String
    Dim LineArrSz As Integer
    Dim FaceArr() As String
    Dim FaceArrSz As Integer
    Dim FaceVertexIdx As Integer
    
    ObjFileNo = FreeFile

    Open ObjPath For Input As ObjFileNo
    Do Until EOF(ObjFileNo)
        Line Input #ObjFileNo, Line
        
        If "mtllib" = left(Line, 6) Then
            LineArrSz = GLStrSplit(LineArr, Line, " ")
            'TODO Load LineArrSz(1) as mtllib.
            'MsgBox LineArr(1)
            
        ElseIf "f" = left(Line, 1) Then
            'Parse the line into an array.
            LineArrSz = GLStrSplit(LineArr, Line, " ")
            If 4 > LineArrSz Then frmLog.LogLine "Invalid array sz: " & LineArrSz
            
            ReDim Preserve ObjIn.Faces(ObjIn.FacesSz) As GLFace
                
            'TODO: Handle >3 vertex indexes.
            For FaceVertexIdx = 1 To 3
                FaceArrSz = GLStrSplit(FaceArr, LineArr(FaceVertexIdx), "/")
            
                'Parse vertex index.
                ReDim Preserve ObjIn.Faces(ObjIn.FacesSz).VertexIdx(ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz) As Integer
                ObjIn.Faces(ObjIn.FacesSz).VertexIdx(ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz) = _
                    Val(FaceArr(0)) - 1 'Vertex indexes are 1-indexed in obj format.
                frmLog.LogLine "Face " & ObjIn.FacesSz & _
                    " Vertex " & ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz & _
                    ": " & ObjIn.Faces(ObjIn.FacesSz).VertexIdx(ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz)
                ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz = ObjIn.Faces(ObjIn.FacesSz).VertexIdxSz + 1
                    
                'Parser normal index.
                If 3 = FaceArrSz Then
                    ReDim Preserve ObjIn.Faces(ObjIn.FacesSz).VNormalIdx(ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz) As Integer
                    ObjIn.Faces(ObjIn.FacesSz).VNormalIdx(ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz) = _
                        Val(FaceArr(2)) - 1 'Vertex indexes are 1-indexed in obj format.
                    frmLog.LogLine "Face " & ObjIn.FacesSz & _
                        " Normal " & ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz & _
                        ": " & ObjIn.Faces(ObjIn.FacesSz).VNormalIdx(ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz)
                    ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz = ObjIn.Faces(ObjIn.FacesSz).VNormalIdxSz + 1
                End If
                
                'TODO: Parse normal/texture indexes.
                    
                
            Next FaceVertexIdx
            
            'Increment vertex count.
            ObjIn.FacesSz = ObjIn.FacesSz + 1
        
        ElseIf "vn" = left(Line, 2) Then
            'Parse the line into an array.
            LineArrSz = GLStrSplit(LineArr, Line, " ")
            If 4 > LineArrSz Then frmLog.LogLine "Invalid array sz: " & LineArrSz
            
            'Prepare and assign vertices.
            ReDim Preserve ObjIn.VNormals(ObjIn.VNormalsSz) As GLVertex
            ObjIn.VNormals(ObjIn.VNormalsSz).x = Val(LineArr(1))
            ObjIn.VNormals(ObjIn.VNormalsSz).y = Val(LineArr(2))
            ObjIn.VNormals(ObjIn.VNormalsSz).z = Val(LineArr(3))
            frmLog.LogLine "VNormal: " & ObjIn.VNormals(ObjIn.VNormalsSz).x & _
                ", " & ObjIn.VNormals(ObjIn.VNormalsSz).y & _
                ", " & ObjIn.VNormals(ObjIn.VNormalsSz).z
            
            'Increment vertex count.
            ObjIn.VNormalsSz = ObjIn.VNormalsSz + 1
            
        ElseIf "vt" = left(Line, 2) Then
            'Parse the line into an array.
            LineArrSz = GLStrSplit(LineArr, Line, " ")
            If 3 > LineArrSz Then frmLog.LogLine "Invalid array sz: " & LineArrSz
            
            'Prepare and assign vertices.
            ReDim Preserve ObjIn.VTextures(ObjIn.VTexturesSz) As GLVTexture
            ObjIn.VTextures(ObjIn.VTexturesSz).u = Val(LineArr(1))
            ObjIn.VTextures(ObjIn.VTexturesSz).V = Val(LineArr(2))
            If 4 = LineArrSz Then ObjIn.VTextures(ObjIn.VTexturesSz).w = Val(LineArr(3))
            frmLog.LogLine "VTexture: " & ObjIn.VTextures(ObjIn.VTexturesSz).u & _
                ", " & ObjIn.VTextures(ObjIn.VTexturesSz).V & _
                ", " & ObjIn.VTextures(ObjIn.VTexturesSz).w
            
            'Increment vertex count.
            ObjIn.VTexturesSz = ObjIn.VTexturesSz + 1
        
        ElseIf "v" = left(Line, 1) Then
            'Parse the line into an array.
            LineArrSz = GLStrSplit(LineArr, Line, " ")
            If 4 > LineArrSz Then frmLog.LogLine "Invalid array sz: " & LineArrSz
            
            'Prepare and assign vertices.
            ReDim Preserve ObjIn.Vertices(ObjIn.VerticesSz) As GLVertex
            ObjIn.Vertices(ObjIn.VerticesSz).x = Val(LineArr(1))
            ObjIn.Vertices(ObjIn.VerticesSz).y = Val(LineArr(2))
            ObjIn.Vertices(ObjIn.VerticesSz).z = Val(LineArr(3))
            frmLog.LogLine "Vertex: " & ObjIn.Vertices(ObjIn.VerticesSz).x & ", " & ObjIn.Vertices(ObjIn.VerticesSz).y & ", " & ObjIn.Vertices(ObjIn.VerticesSz).z
            
            'Increment vertex count.
            ObjIn.VerticesSz = ObjIn.VerticesSz + 1
            
        Else
            'frmLog.LogLine Line
        End If
        
        DoEvents
    Loop
    Close ObjFileNo
    
End Function

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
        GLShowError = True
    End If
End Function
Public Function GLStrSplit(ByRef StrOut() As String, ByVal StrIn As String, ByVal Token As String) As Integer

    Dim WordsOut As Integer
    Dim CharIdx As Integer

    ReDim StrOut(0) As String
    For CharIdx = 1 To Len(StrIn)
        If right(left(StrIn, CharIdx), 1) = Token Then
            WordsOut = WordsOut + 1
            
            'Realloc but preserve what's parsed so far!
            ReDim Preserve StrOut(WordsOut) As String
            StrOut(WordsOut) = ""
        Else
            StrOut(WordsOut) = StrOut(WordsOut) & right(left(StrIn, CharIdx), 1)
        End If
    Next
    
    GLStrSplit = WordsOut + 1

End Function
