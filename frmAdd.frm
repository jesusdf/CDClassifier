VERSION 5.00
Begin VB.Form frmAdd 
   BackColor       =   &H00400000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4470
   ControlBox      =   0   'False
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbDrives 
      Height          =   315
      Left            =   990
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "Lista de Unidades de CD"
      Top             =   330
      Width           =   1605
   End
   Begin VB.CommandButton cmdClose 
      Height          =   645
      Left            =   3600
      Picture         =   "frmAdd.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar la ventana"
      Top             =   150
      Width           =   735
   End
   Begin VB.CommandButton cmdAnalize 
      Height          =   645
      Left            =   2760
      Picture         =   "frmAdd.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Analizar el CD"
      Top             =   150
      Width           =   705
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   645
      Left            =   150
      Picture         =   "frmAdd.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Abrir la Bandeja"
      Top             =   150
      Width           =   705
   End
   Begin VB.Timer tmrIntermitente 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   240
      Top             =   210
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   2970
      TabIndex        =   12
      Top             =   240
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tamaño del CD:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   14
      Top             =   1620
      Width           =   1155
   End
   Begin VB.Label lblSize 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1620
      TabIndex        =   13
      Top             =   1620
      Width           =   2445
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   11
      Top             =   2370
      Width           =   900
   End
   Begin VB.Label lblProcesando 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "C:\"
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2670
      Width           =   3735
   End
   Begin VB.Label lblEstado 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Analizando"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   1500
      TabIndex        =   9
      Top             =   1890
      Width           =   2565
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado Actual:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   8
      Top             =   1890
      Width           =   1035
   End
   Begin VB.Label lblSerie 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1800
      TabIndex        =   7
      Top             =   1350
      Width           =   2265
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº de Serie del CD:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   6
      Top             =   1350
      Width           =   1380
   End
   Begin VB.Label lblEtiqueta 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   1590
      TabIndex        =   5
      Top             =   1080
      Width           =   2475
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Etiqueta del CD:"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   390
      TabIndex        =   4
      Top             =   1080
      Width           =   1155
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EsVisible As Boolean

Private Sub CargaUnidadesEnCombo()

On Error Resume Next

Dim All As String
Dim pos As Long
Dim JustOne As String
Dim MiEtiqueta As String
Dim i As Long

cmbDrives.Clear
All = GetDrives
Do
    pos = InStr(All, Chr(0))
    If pos Then
        JustOne = UCase(Left(All, pos - 1))
        If EsUnaUnidadDeCDROM(JustOne) = True Then
            i = DimeElNumeroDeSerie(Mid(JustOne, 1, 2), MiEtiqueta)
            If MiEtiqueta = vbNullString And i = 0 Then MiEtiqueta = "Vacío"
            cmbDrives.AddItem JustOne & " (" & MiEtiqueta & ")"
        End If
        All = Mid(All, pos + 1, Len(All))
    End If
    DoEvents
Loop Until All = ""

If cmbDrives.ListCount > 0 Then cmbDrives.ListIndex = 0

End Sub

Private Function AñadeDirectoriosYArchivos(Path As String, ParentID As Long) As Long
On Error Resume Next
Dim Directorio As String
Dim Archivo As String
Dim ParentIDLocal As Long
Dim i As Long
If Right(Path, 1) <> "\" Then Path = Path & "\"
lblProcesando.Caption = Path
DoEvents
Dir1.Path = Path
Archivo = Dir(Path & "*.*", vbArchive + vbNormal + vbHidden + vbReadOnly + vbSystem)
While Archivo <> ""
    tabla2.AddNew
    tabla2.Fields("file") = Archivo
    tabla2.Fields("dirid") = ParentID
    tabla2.Fields("cdid") = CDSerial
    tabla2.Fields("comentario") = ""
    tabla2.Update
    Archivo = Dir
Wend
i = 0
While i <= Dir1.ListCount - 1
    Directorio = DirActual(Dir1.List(i))
    Tabla.AddNew
    Tabla.Fields("dir") = Directorio
    Tabla.Fields("parentid") = ParentID
    Tabla.Fields("cdid") = CDSerial
    ParentIDLocal = Tabla.Fields("dirid")
    Tabla.Update
    If Dir(Path & Directorio, vbDirectory) <> "" Then AñadeDirectoriosYArchivos Path & Directorio, ParentIDLocal
    Dir1.Path = Path
    i = i + 1
Wend
End Function

Private Sub cmdAnalize_Click()
On Error GoTo Hell

Dim i As Long
Dim Path As String
Dim Directorio As String
Dim raizid As Long

CDRom = Mid(cmbDrives.List(cmbDrives.ListIndex), 1, 3)

CDSerial = DimeElNumeroDeSerie(CDRom, CDLabel)

If CDSerial = 0 Then
    MsgBox "La unidad está vacía, por favor introduce un CD primero.", vbExclamation, "Error"
    cmdOpen.SetFocus
    Exit Sub
End If

Dir1.Path = CDRom
Dir1.Refresh

cmbDrives.Enabled = False
cmdOpen.Enabled = False
cmdAnalize.Enabled = False
cmdClose.Enabled = False

lblEtiqueta.Caption = CDLabel
lblSerie.Caption = Hex(CDSerial)
lblProcesando.Caption = CDRom
lblEstado.Caption = "Analizando"
lblSize.Caption = Trim(Str(DiskSize(CDRom) / 1024 \ 1024)) & " MB"

While Me.Height < (lblProcesando.Top + lblProcesando.Height + 240)
    Me.Height = Me.Height + 24
    Me.Top = Me.Top - 12
    DoEvents
Wend

tmrIntermitente.Enabled = True

' Añado el CD al catálogo

Set Tabla = bd.OpenRecordset("tcds", dbOpenDynaset)

Tabla.AddNew
Tabla.Fields("cdid") = CDSerial
Tabla.Fields("cdlabel") = CDLabel
Tabla.Fields("totalsize") = Trim(Str(DiskSize(CDRom)))
Tabla.Update
Tabla.Close

Set Tabla = bd.OpenRecordset("tdirs", dbOpenDynaset)
Set tabla2 = bd.OpenRecordset("tfiles", dbOpenDynaset)
AñadeDirectoriosYArchivos CDRom, -1
Tabla.Close
tabla2.Close

lblEstado.ForeColor = vbGreen
lblEstado.Caption = "Finalizado"
lblProcesando.Caption = CDRom
tmrIntermitente.Enabled = False
lblEstado.Visible = True
cmbDrives.Enabled = True
cmdOpen.Enabled = True
cmdAnalize.Enabled = True
cmdClose.Enabled = True
cmdClose.SetFocus

Exit Sub

Hell:
    Select Case Err.Number
        Case 3022
            If MsgBox("Ya existe un registro del CD-ROM '" & CDLabel & " - " & Hex(CDSerial) & "' en la base de datos." & vbCrLf & "¿Deseas borrarlo y volver a analizarlo?", vbQuestion + vbYesNo, "¿Actualizar?") = vbYes Then
                Tabla.Close
                Set Tabla = bd.OpenRecordset("tcds", dbOpenDynaset)
                Tabla.FindFirst "cdid=" & CDSerial
                Tabla.Delete
                cmdAnalize_Click
                Exit Sub
            Else
                Tabla.Close
                lblEstado.Caption = "Cancelado"
                tmrIntermitente.Enabled = False
                lblEstado.Visible = True
                cmbDrives.Enabled = True
                cmdOpen.Enabled = True
                cmdAnalize.Enabled = True
                cmdClose.Enabled = True
                Exit Sub
            End If
        Case Else
            MsgBox Err.Description, vbExclamation, "Error"
            lblEstado.Visible = True
            cmbDrives.Enabled = True
            cmdOpen.Enabled = True
            cmdAnalize.Enabled = True
            cmdClose.Enabled = True
    End Select
    Resume Next
End Sub

Private Sub cmdOpen_Click()
Dim lRet As Long
Dim openParams As MCI_OPEN_PARMS

    openParams.lpstrDeviceType = "cdaudio"
    openParams.lpstrElementName = Mid(cmbDrives.List(cmbDrives.ListIndex), 1, 2) & Chr(0)
    lRet = mciSendCommand(0, MCI_OPEN, MCI_OPEN_TYPE Or MCI_OPEN_SHAREABLE Or MCI_OPEN_ELEMENT, openParams)
    If lRet = 0 Then
        lRet = mciSendCommand(openParams.wDeviceID, MCI_SET, MCI_SET_DOOR_OPEN, ByVal 0&)
    End If
    MsgBox "Pulsa [Aceptar] para cerrar la bandeja.", vbInformation, "CD-ROM Abierto"
    lRet = mciSendCommand(openParams.wDeviceID, MCI_SET, MCI_SET_DOOR_CLOSED, ByVal 0&)
    mciSendCommand openParams.wDeviceID, MCI_CLOSE, MCI_WAIT, ByVal 0&

'SendMCIString "set cdaudio door open"
'MsgBox "Pulsa [Aceptar] para cerrar la bandeja.", vbInformation, "CD-ROM Abierto"
'SendMCIString "set cdaudio door closed"

CargaUnidadesEnCombo

End Sub

Private Sub cmdClose_Click()
If lblEstado.ForeColor = vbGreen Then
    While Me.Height > (cmdOpen.Top + cmdOpen.Height + 240)
        Me.Height = Me.Height - 24
        Me.Top = Me.Top + 12
        DoEvents
    Wend
    Espera 0.5
End If
tmrIntermitente.Enabled = False
Unload Me
End Sub

Private Sub Form_Load()
CargaUnidadesEnCombo
EsVisible = True
End Sub

Private Sub tmrIntermitente_Timer()
If EsVisible = True Then
    EsVisible = False
    lblEstado.Visible = EsVisible
Else
    EsVisible = True
    lblEstado.Visible = EsVisible
End If
End Sub
