VERSION 4.00
Begin VB.Form frmLights 
   Caption         =   "Lighting"
   ClientHeight    =   4245
   ClientLeft      =   7770
   ClientTop       =   3960
   ClientWidth     =   7605
   Height          =   4650
   Icon            =   "lights.frx":0000
   Left            =   7710
   LinkTopic       =   "Form1"
   ScaleHeight     =   4245
   ScaleWidth      =   7605
   Top             =   3615
   Visible         =   0   'False
   Width           =   7725
   Begin VB.CheckBox ChkLight0 
      Caption         =   "Enable Light 0"
      Height          =   255
      Left            =   3960
      TabIndex        =   0
      Top             =   120
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Z"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   3720
      Width           =   255
   End
   Begin VB.Label Label6 
      Caption         =   "Y"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "X"
      Height          =   255
      Left            =   240
      TabIndex        =   21
      Top             =   2760
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "W"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   2040
      Width           =   255
   End
   Begin VB.Label Label3 
      Caption         =   "Z"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1560
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "Y"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   255
   End
   Begin VB.Label Label1 
      Caption         =   "X"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "Position:"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Direction:"
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Lbl0DZ 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   14
      Top             =   3600
      Width           =   615
   End
   Begin VB.Label Lbl0DY 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   13
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Lbl0DX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   12
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label Lbl0PW 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Lbl0PZ 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   10
      Top             =   1440
      Width           =   615
   End
   Begin VB.Label Lbl0PY 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Lbl0PX 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   480
      Width           =   615
   End
   Begin ComctlLib.Slider Slider0DX 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0DY 
      Height          =   375
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0DZ 
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   3600
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0W 
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   1920
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0Z 
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1440
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0Y 
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   960
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
   Begin ComctlLib.Slider Slider0X 
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   480
      Width           =   6015
      _Version        =   65536
      _ExtentX        =   10610
      _ExtentY        =   661
      _StockProps     =   64
      Min             =   -10
   End
End
Attribute VB_Name = "frmLights"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
