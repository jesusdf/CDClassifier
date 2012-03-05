VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSearch 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   8700
   ControlBox      =   0   'False
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExecute 
      Height          =   645
      Left            =   6930
      Picture         =   "frmSearch.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Abrir el archivo"
      Top             =   2820
      Width           =   735
   End
   Begin VB.CommandButton cmdClose 
      Height          =   645
      Left            =   7800
      Picture         =   "frmSearch.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cerrar la ventana"
      Top             =   2820
      Width           =   705
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2505
      Left            =   180
      TabIndex        =   0
      ToolTipText     =   "Si haces doble click sobre una fila verás un mensaje con todo el texto en el caso en el que se vea cortado"
      Top             =   180
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   4419
      _Version        =   393216
      Rows            =   0
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   14737632
      BackColorFixed  =   14737632
      BackColorSel    =   12648384
      AllowBigSelection=   0   'False
      GridLines       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2580
      TabIndex        =   4
      Top             =   3060
      Width           =   2145
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de resultados de la búsqueda:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   210
      TabIndex        =   3
      Top             =   3060
      Width           =   2355
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function NombreCD(ByVal CDID As Long) As String
Dim tablalocal As Recordset
Dim Nombre As String
Set tablalocal = bd.OpenRecordset("SELECT * FROM tcds WHERE cdid=" & CDID)
If Not tablalocal.EOF Then
    Nombre = tablalocal.Fields("cdlabel") & " - " & Hex(CDID)
End If
tablalocal.Close
NombreCD = Nombre
End Function

Private Function CreaPath(ByVal ID As Long, ByVal Tipo As String, ByVal CDID As Long) As String
Dim tablalocal As Recordset
Dim mitabla As String
Dim campo As String
Dim Path As String
If Tipo <> "Directorio" Then
    ' Si es un archivo o directorio, primero hay que ver cual es su directorio padre
    Set tablalocal = bd.OpenRecordset("SELECT * FROM tfiles WHERE cdid=" & CDID & " and fileid=" & ID)
    ID = tablalocal.Fields("dirid")
    Path = tablalocal.Fields("file")
    If ID = -1 Then
        Path = "X:\" & Path
        CreaPath = Path
        Exit Function
    End If
    tablalocal.Close
End If
While ID <> -1
    Set tablalocal = bd.OpenRecordset("SELECT * FROM tdirs WHERE cdid=" & CDID & " and dirid=" & ID)
    If Not tablalocal.EOF Then
        Path = tablalocal.Fields("dir") & "\" & Path
        ID = tablalocal.Fields("parentid")
    End If
    tablalocal.Close
Wend
Path = "X:\" & Path
CreaPath = Path
End Function

Private Sub AñadeFila(ByVal CD As String, ByVal Tipo As String, ByVal Ruta As String)
Grid.AddItem Grid.Rows + 1
Grid.Row = Grid.Rows - 1
Grid.Col = 1
Grid.Text = CD
Grid.CellAlignment = flexAlignCenterCenter
Grid.Col = 2
Grid.Text = Tipo
Grid.CellAlignment = flexAlignCenterCenter
Grid.Col = 3
Grid.Text = Ruta
Grid.CellAlignment = flexAlignCenterCenter
End Sub

Private Sub Busca(Tabla As String, campo As String, CampoID As String, Tipo As String)
On Error Resume Next
Dim tablalocal As Recordset
Dim QueBusco As String
Dim AñadidoSQL As String
QueBusco = Filtra(frmView.txtSearch.Text)
If frmView.optActual.Value = True Then AñadidoSQL = " and cdid=" & frmView.cmbCDs.ItemData(frmView.cmbCDs.ListIndex)
Set tablalocal = bd.OpenRecordset("SELECT * FROM " & Tabla & " WHERE " & campo & " like '" & QueBusco & "'" & AñadidoSQL, dbOpenDynaset)
While Not tablalocal.EOF
    AñadeFila NombreCD(tablalocal.Fields("cdid")), Tipo, CreaPath(tablalocal.Fields(CampoID), IIf(Tipo = "Comentario", "Archivo", Tipo), tablalocal.Fields("cdid"))
    tablalocal.MoveNext
Wend
If tablalocal.RecordCount > 0 Then
    Grid.Col = 0
    Grid.Row = 0
    Grid.ColSel = Grid.Cols - 1
    Grid.RowSel = 0
End If
tablalocal.Close
End Sub

Private Sub BuscaArchivo()
If frmView.chkFiles.Value = vbChecked Then
    Busca "tfiles", "file", "fileid", "Archivo"
End If
If frmView.chkDirs.Value = vbChecked Then
    Busca "tdirs", "dir", "dirid", "Directorio"
End If
If frmView.chkComments.Value = vbChecked Then
    Busca "tfiles", "comentario", "fileid", "Comentario"
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdExecute_Click()
On Error Resume Next
Dim CDID As String
Dim All As String
Dim JustOne As String
Dim pos As Long
Dim MiEtiqueta As String
Dim i As Long

Grid.Col = 1
CDID = Grid.Text

All = GetDrives
Do
    pos = InStr(All, Chr(0))
    If pos Then
        JustOne = UCase(Left(All, pos - 1))
        If EsUnaUnidadDeCDROM(JustOne) = True Then
            i = DimeElNumeroDeSerie(Mid(JustOne, 1, 2), MiEtiqueta)
            If MiEtiqueta <> vbNullString And i <> 0 Then
                If MiEtiqueta & " - " & Hex(i) = CDID Then
                    Grid.Col = 3
                    Ejecuta JustOne & Mid(Grid.Text, 3)
                    Grid.Col = 0
                    Exit Sub
                End If
            End If
        End If
        All = Mid(All, pos + 1, Len(All))
    End If
    DoEvents
Loop Until All = ""

Grid.Col = 0
MsgBox "No se ha encontrado el CD correspondiente en ninguna Unidad, introdúcelo y vuelve a intentarlo.", vbExclamation, "Ups"

End Sub

Private Sub Form_Load()
Grid.ColWidth(0) = 500  ' Nº
Grid.ColWidth(1) = 2000 ' CD
Grid.ColWidth(2) = 1500 ' Tipo
Grid.ColWidth(3) = 4000 ' Ruta
BuscaArchivo
lblCount.Caption = Grid.Rows
End Sub

Private Sub Grid_DblClick()
On Error GoTo Hell
Dim Texto As String
Dim i As Integer
Grid.Col = 0
Texto = Grid.Text
For i = 1 To 3
    Grid.Col = i
    Texto = Texto & " | " & Grid.Text
Next i
Grid.Col = 0
MsgBox Texto, vbInformation, "Ampliación"
Exit Sub
Hell:
    Exit Sub
End Sub
