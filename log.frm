VERSION 4.00
Begin VB.Form frmLog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Log View"
   ClientHeight    =   2295
   ClientLeft      =   1665
   ClientTop       =   8535
   ClientWidth     =   6735
   Height          =   2700
   Icon            =   "log.frx":0000
   Left            =   1605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6735
   Top             =   8190
   Width           =   6855
   Begin VB.CheckBox ChkAutoScroll 
      Caption         =   "&Autoscroll"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.TextBox TextLog 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   6495
   End
End
Attribute VB_Name = "frmLog"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

'
Public Sub LogLine(ByVal Line As String)
    If Not frmLog.Visible Then frmLog.Show
    
    TextLog.Text = frmLog.TextLog.Text & Chr(13) & Chr(10) & Line
    'textlog.
End Sub

