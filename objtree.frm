VERSION 4.00
Begin VB.Form frmObjTree 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tree View"
   ClientHeight    =   5895
   ClientLeft      =   7770
   ClientTop       =   2310
   ClientWidth     =   4815
   Height          =   6300
   Icon            =   "objtree.frx":0000
   Left            =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   4815
   Top             =   1965
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
      NumImages       =   2
      i1              =   "objtree.frx":0442
      i2              =   "objtree.frx":05F9
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


