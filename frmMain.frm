VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CDClassifier v1.0"
   ClientHeight    =   4470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5985
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBackground 
      BorderStyle     =   0  'None
      Height          =   4485
      Left            =   0
      Picture         =   "frmMain.frx":08CA
      ScaleHeight     =   4485
      ScaleWidth      =   6015
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.CommandButton cmdAdd 
         Height          =   855
         Left            =   360
         Picture         =   "frmMain.frx":446B
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CommandButton cmdView 
         Height          =   855
         Left            =   360
         Picture         =   "frmMain.frx":5335
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2370
         Width           =   1095
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   855
         Left            =   360
         Picture         =   "frmMain.frx":61FF
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblInfo1 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00E0E0E0&
         Height          =   825
         Left            =   2310
         TabIndex        =   4
         Top             =   1500
         Width           =   3465
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblInfo2 
         BackStyle       =   0  'Transparent
         Height          =   675
         Left            =   2340
         TabIndex        =   5
         Top             =   1530
         Width           =   3465
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim texto(3) As String

Private Sub cmdAdd_Click()
frmAdd.Show vbModal, Me
End Sub

Private Sub cmdAdd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblInfo1.Caption <> texto(0) Then
    lblInfo1.Caption = texto(0)
    lblInfo2.Caption = lblInfo1.Caption
End If
End Sub

Private Sub cmdSearch_Click()
Unload Me
End Sub

Private Sub cmdSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblInfo1.Caption <> texto(2) Then
    lblInfo1.Caption = texto(2)
    lblInfo2.Caption = lblInfo1.Caption
End If
End Sub

Private Sub cmdView_Click()
frmView.Show vbModal, Me
End Sub

Private Sub cmdView_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblInfo1.Caption <> texto(1) Then
    lblInfo1.Caption = texto(1)
    lblInfo2.Caption = lblInfo1.Caption
End If
End Sub

Private Sub Form_Load()
texto(0) = "&Añadir CD al catálogo"
texto(1) = "&Inspeccionar el Catálogo"
texto(2) = "&Salir del programa"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lblInfo1.Caption <> vbNullString Then
    lblInfo1.Caption = vbNullString
    lblInfo2.Caption = lblInfo1.Caption
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
bd.Close
End
End Sub

Private Sub picBackground_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Form_MouseMove Button, Shift, X, Y
End Sub
