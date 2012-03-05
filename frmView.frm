VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmView 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9375
   ControlBox      =   0   'False
   Icon            =   "frmView.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExport 
      Height          =   645
      Left            =   7590
      Picture         =   "frmView.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Exportar los contenidos del CD a formato HTML"
      Top             =   90
      Width           =   705
   End
   Begin VB.CommandButton cmdCompactAndRepair 
      Height          =   645
      Left            =   6750
      Picture         =   "frmView.frx":0E54
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Reparar y compactar el catálogo"
      Top             =   90
      Width           =   705
   End
   Begin VB.CommandButton cmdDelete 
      Height          =   645
      Left            =   2790
      Picture         =   "frmView.frx":171E
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Borrar el CD del catálogo"
      Top             =   90
      Width           =   705
   End
   Begin MSComctlLib.ImageList ImageList_File 
      Left            =   780
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":1FE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdClose 
      Height          =   645
      Left            =   8430
      Picture         =   "frmView.frx":28C2
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cerrar la ventana"
      Top             =   90
      Width           =   705
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00400000&
      Caption         =   "Comentario del archivo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   4680
      TabIndex        =   8
      Top             =   5820
      Width           =   4455
      Begin VB.CommandButton cmdEdit 
         Height          =   675
         Left            =   3630
         Picture         =   "frmView.frx":318C
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Editar el comentario"
         Top             =   240
         Width           =   705
      End
      Begin VB.TextBox txtComment 
         Height          =   1065
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   240
         Width           =   3375
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00400000&
      Caption         =   "Buscar:"
      ForeColor       =   &H00FFFFFF&
      Height          =   1485
      Left            =   180
      TabIndex        =   3
      Top             =   5820
      Width           =   4305
      Begin VB.OptionButton optAll 
         BackColor       =   &H00400000&
         Caption         =   "En todo el catálogo"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   2190
         TabIndex        =   14
         Top             =   1110
         Width           =   1725
      End
      Begin VB.OptionButton optActual 
         BackColor       =   &H00400000&
         Caption         =   "En el CD actual"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   330
         TabIndex        =   13
         Top             =   1110
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox chkComments 
         BackColor       =   &H00400000&
         Caption         =   "Comentarios"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2190
         TabIndex        =   12
         Top             =   660
         Width           =   1215
      End
      Begin VB.CheckBox chkDirs 
         BackColor       =   &H00400000&
         Caption         =   "Directorios"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1110
         TabIndex        =   7
         Top             =   660
         Width           =   1095
      End
      Begin VB.TextBox txtSearch 
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Text            =   "Escribe aquí tu búsqueda"
         ToolTipText     =   "Recuerda que puedes utilizar comodines: * ?"
         Top             =   240
         Width           =   3165
      End
      Begin VB.CheckBox chkFiles 
         BackColor       =   &H00400000&
         Caption         =   "Archivos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   150
         TabIndex        =   5
         Top             =   660
         Value           =   1  'Checked
         Width           =   945
      End
      Begin VB.CommandButton cmdSearch 
         Height          =   675
         Left            =   3480
         Picture         =   "frmView.frx":3A56
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Buscar"
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.ComboBox cmbCDs 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2445
   End
   Begin MSComctlLib.ListView ListView 
      Height          =   4815
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   8493
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList_File"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   210
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":4320
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":48BA
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":4E54
            Key             =   "closedold"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":4FAE
            Key             =   "openold"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmView.frx":5108
            Key             =   "cd"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView 
      Height          =   4815
      Left            =   180
      TabIndex        =   0
      Top             =   840
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   8493
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      ImageList       =   "ImageList"
      Appearance      =   1
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   5070
      TabIndex        =   19
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label lblSerial 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   5070
      TabIndex        =   18
      Top             =   150
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño del CD:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3630
      TabIndex        =   17
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie del CD:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3630
      TabIndex        =   16
      Top             =   150
      Width           =   1380
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaCDs()
On Error Resume Next
Dim tablalocal As Recordset
cmbCDs.Clear
TreeView.Nodes.Clear
ListView.ListItems.Clear
Set tablalocal = bd.OpenRecordset("SELECT * FROM tcds ORDER BY cdlabel", dbOpenDynaset)
While Not tablalocal.EOF
    cmbCDs.AddItem tablalocal.Fields("cdlabel") & " - " & Hex(tablalocal.Fields("cdid"))
    cmbCDs.ItemData(cmbCDs.NewIndex) = tablalocal.Fields("cdid")
    tablalocal.MoveNext
Wend
tablalocal.Close
If cmbCDs.ListCount > 0 Then cmbCDs.ListIndex = 0
End Sub

Private Function ListaArchivos(ParentID As Long, Profundidad As Long) As String
On Error GoTo Hell
Dim i As Long
Dim Espacios As String
Dim mibuffer As String
Dim tablalocal As Recordset

For i = 1 To Profundidad
    Espacios = Espacios & "&nbsp;&nbsp;&nbsp;&nbsp;"
Next i
Set tablalocal = bd.OpenRecordset("SELECT * FROM tfiles WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex) & " and dirid=" & ParentID, dbOpenDynaset)
While Not tablalocal.EOF
    mibuffer = mibuffer & Espacios & "<img border=""0"" align=""absmiddle"" src=""file.gif""><font face=""Courier New"" size=""2"">   " & tablalocal.Fields("file") & "</font><br>" & vbCrLf
    tablalocal.MoveNext
Wend
tablalocal.Close
ListaArchivos = mibuffer
Exit Function
Hell:
    Exit Function
End Function

Private Function CreaListado(ParentID As Long, Profundidad As Long) As String
On Error GoTo Hell
Dim i As Long
Dim Espacios As String
Dim mibuffer As String
Dim tablalocal As Recordset

For i = 1 To Profundidad
    Espacios = Espacios & "&nbsp;&nbsp;&nbsp;&nbsp;"
Next i
Set tablalocal = bd.OpenRecordset("SELECT * FROM tdirs WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex) & " and parentid=" & ParentID, dbOpenDynaset)
While Not tablalocal.EOF
    mibuffer = mibuffer & Espacios & "<img border=""0"" align=""absmiddle"" src=""dir.gif""><font face=""Courier New"" size=""2"">   " & tablalocal.Fields("dir") & "</font><br>" & vbCrLf
    mibuffer = mibuffer & CreaListado(tablalocal.Fields("dirid"), Profundidad + 1)
    mibuffer = mibuffer & ListaArchivos(tablalocal.Fields("dirid"), Profundidad + 1)
    tablalocal.MoveNext
Wend
tablalocal.Close
CreaListado = mibuffer
Exit Function
Hell:
    Exit Function
End Function

Private Sub BuscaSubDirectorios(ParentID As Long)
On Error Resume Next
Dim tablalocal As Recordset
Set tablalocal = bd.OpenRecordset("SELECT * FROM tdirs WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex) & " and parentid=" & ParentID, dbOpenDynaset)
While Not tablalocal.EOF
    AñadeDirectorio tablalocal.Fields("dir"), tablalocal.Fields("dirid"), ParentID
    BuscaSubDirectorios tablalocal.Fields("dirid")
    tablalocal.MoveNext
Wend
tablalocal.Close
End Sub

Private Sub BuscaDirectorios()
On Error Resume Next
tabla2.FindFirst "parentid=-1"
While Not tabla2.NoMatch
    AñadeDirectorio tabla2.Fields("dir"), tabla2.Fields("dirid"), -1
    BuscaSubDirectorios tabla2.Fields("dirid")
    tabla2.FindNext "parentid=-1"
Wend
BuscaArchivos -1
tabla2.MoveFirst
End Sub

Private Sub BuscaArchivos(ParentID As Long)
On Error Resume Next
ListView.ListItems.Clear
tabla3.FindFirst "dirid=" & ParentID
While Not tabla3.NoMatch
    AñadeArchivo tabla3.Fields("file"), tabla3.Fields("fileid"), tabla3.Fields("comentario")
    tabla3.FindNext "dirid=" & ParentID
Wend
tabla3.MoveFirst
End Sub

Private Sub AñadeDirectorioRaiz(Directorio As String, DirID As Long)
On Error Resume Next
Dim Nodo As Node
Set Nodo = TreeView.Nodes.Add(, , "C" & DirID, Directorio, "cd", "cd")
Nodo.Expanded = True
End Sub

Private Sub AñadeDirectorio(Directorio As String, DirID As Long, Parent As Long)
On Error Resume Next
'.Nodes.Clear
Dim Nodo As Node
Set Nodo = TreeView.Nodes.Add("C" & Parent, tvwChild, "C" & DirID, Directorio, "closed", "open")
End Sub

Private Sub AñadeArchivo(Archivo As String, FileID As Long, Comentario As String)
On Error Resume Next
Dim itmX As ListItem
Set itmX = ListView.ListItems.Add()
itmX.Tag = Comentario
itmX.Icon = 1
itmX.Key = "F" & FileID
itmX.Text = Archivo
End Sub

Private Sub cmbCDs_Click()
On Error Resume Next
TreeView.Nodes.Clear
ListView.ListItems.Clear
Tabla.FindFirst "cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex)
If Not Tabla.NoMatch Then
    lblSerial.Caption = Hex(cmbCDs.ItemData(cmbCDs.ListIndex))
    lblSize.Caption = Trim(Str(Tabla.Fields("totalsize") / 1024 \ 1024)) & " MB"
    tabla2.Close
    tabla3.Close
    AñadeDirectorioRaiz "X:\", -1
    Set tabla2 = bd.OpenRecordset("SELECT * FROM tdirs WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex), dbOpenDynaset)
    Set tabla3 = bd.OpenRecordset("SELECT * FROM tfiles WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex), dbOpenDynaset)
    BuscaDirectorios
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCompactAndRepair_Click()
On Error GoTo Hell
Dim bdtemporal As String
Dim bdactual As String
bdtemporal = GetTempDir & "~CD_Repair_Compact.mdb"
bdactual = AppPath & "CDs.mdb"
Tabla.Close
tabla2.Close
tabla3.Close
bd.Close
DoEvents
RepairDatabase bdactual
CompactDatabase bdactual, bdtemporal
Kill bdactual
FileCopy bdtemporal, bdactual
Kill bdtemporal
Set bd = OpenDatabase(bdactual)
Set Tabla = bd.OpenRecordset("tcds", dbOpenDynaset)
Set tabla2 = bd.OpenRecordset("SELECT * FROM tdirs WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex), dbOpenDynaset)
Set tabla3 = bd.OpenRecordset("SELECT * FROM tfiles WHERE cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex), dbOpenDynaset)
MsgBox "El catálogo ha sido Reparado y Compactado correctamente. :)", vbInformation, "Ok!"
Exit Sub
Hell:
    'If Err.Number <> 3251 Then
    '    MsgBox Err.Description, vbCritical, "Error"
    'End If
    Resume Next
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
If MsgBox("¿Estas seguro de que quieres eliminar este CD del catálogo?", vbQuestion + vbYesNo, "Confirma") = vbYes Then
    Tabla.FindFirst "cdid=" & cmbCDs.ItemData(cmbCDs.ListIndex)
    If Not Tabla.NoMatch Then
        Tabla.Delete
        CargaCDs
    End If
End If
End Sub

Private Sub cmdEdit_Click()
On Error Resume Next
Dim QuerySQL As String
QuerySQL = "UPDATE tfiles SET tfiles.comentario = '" & Filtra(txtComment.Text) & "' WHERE (((tfiles.fileid)=" & Mid(ListView.SelectedItem.Key, 2) & "));"
bd.Execute QuerySQL
ListView.SelectedItem.Tag = Filtra(txtComment.Text)
MsgBox "Comentario actualizado. :)", vbInformation, "Ok!"
End Sub

Private Sub cmdSearch_Click()
On Error Resume Next
If chkFiles.Value = vbUnchecked And chkDirs.Value = vbUnchecked And chkComments.Value = vbUnchecked Then
    MsgBox "Debes elegir al menos una opción donde buscar.", vbExclamation, "Error"
    chkFiles.SetFocus
    Exit Sub
End If
frmSearch.Show vbModal
End Sub

Private Sub cmdExport_Click()
On Error Resume Next
Dim buffer As String
MkDir AppPath & "HTML\"
buffer = "<html>" & vbCrLf & "<head>" & vbCrLf
buffer = buffer & "<title>Árbol de directorios del CD: " & cmbCDs.List(cmbCDs.ListIndex) & "</title>"
buffer = buffer & "</head>" & vbCrLf & "<body>" & vbCrLf & "<h1>" & cmbCDs.List(cmbCDs.ListIndex) & "</h1>"
buffer = buffer & "<img border=""0"" align=""absmiddle"" src=""cd.gif""><font face=""Courier New"" size=""2"">   X:\</font><br>" & vbCrLf
buffer = buffer & CreaListado(-1, 1)
buffer = buffer & ListaArchivos(-1, 0) & vbCrLf
buffer = buffer & "<br><br><font face=""Courier New"" size=""2"">Creado con CDClassifier v1.0</font></body>" & vbCrLf & "</html>"
Kill AppPath & "HTML\" & cmbCDs.List(cmbCDs.ListIndex) & ".html"
Open AppPath & "HTML\" & cmbCDs.List(cmbCDs.ListIndex) & ".html" For Binary As #1
    Put 1, , buffer
Close #1
Ejecuta AppPath & "HTML\" & cmbCDs.List(cmbCDs.ListIndex) & ".html"
End Sub

Private Sub Form_Load()
On Error Resume Next
Set Tabla = bd.OpenRecordset("tcds", dbOpenDynaset)
CargaCDs
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Tabla.Close
tabla2.Close
tabla3.Close
End Sub

Private Sub ListView_ItemClick(ByVal item As MSComctlLib.ListItem)
txtComment.Text = item.Tag
End Sub

Private Sub TreeView_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
BuscaArchivos CLng(Mid(Node.Key, 2))
End Sub

Private Sub txtSearch_GotFocus()
On Error Resume Next
txtSearch.SelStart = 0
txtSearch.SelLength = Len(txtSearch.Text)
End Sub
