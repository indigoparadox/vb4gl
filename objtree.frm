VERSION 4.00
Begin VB.Form frmObjTree 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   2610
   ClientTop       =   1830
   ClientWidth     =   4815
   Height          =   6315
   Left            =   2550
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   4815
   Top             =   1485
   Width           =   4935
   Begin ComctlLib.ImageList ImageListObj 
      Left            =   120
      Top             =   5160
      _Version        =   65536
      _ExtentX        =   1005
      _ExtentY        =   1005
      _StockProps     =   1
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      NumImages       =   1
      i1              =   "objtree.frx":0000
   End
   Begin ComctlLib.TreeView TreeViewObj 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      _Version        =   65536
      _ExtentX        =   8070
      _ExtentY        =   9975
      _StockProps     =   196
      Appearance      =   1
      ImageList       =   "ImageListObj"
      PathSeparator   =   "\"
      Style           =   7
   End
End
Attribute VB_Name = "frmObjTree"
Attribute VB_Creatable = False
Attribute VB_Exposed = False


