Attribute VB_Name = "Modulo"
Option Explicit

Public Const DRIVE_REMOVABLE = 2
Public Const DRIVE_FIXED = 3
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_CDROM = 5
Public Const DRIVE_RAMDISK = 6
Public Const MCI_OPEN = &H803
Public Const MCI_OPEN_TYPE = &H2000&
Public Const MCI_OPEN_SHAREABLE = &H100&
Public Const MCI_OPEN_ELEMENT = &H200&
Public Const MCI_OPEN_TYPE_ID = &H1000&
Public Const MCI_SET = &H80D
Public Const MCI_SET_DOOR_OPEN = &H100&
Public Const MCI_SET_DOOR_CLOSED = &H200&
Public Const MCI_DEVTYPE_CD_AUDIO = 516
Public Const MCI_CLOSE = &H804
Public Const MCI_WAIT = &H2&

Public Type MCI_OPEN_PARMS
    dwCallback As Long
    wDeviceID As Long
    lpstrDeviceType As String
    lpstrElementName As String
    lpstrAlias As String
End Type

Public Declare Function mciSendCommand Lib "winmm.dll" Alias "mciSendCommandA" (ByVal wDeviceID As Long, ByVal uMessage As Long, ByVal dwParam1 As Long, ByRef dwParam2 As Any) As Long
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public bd As Database
Public Tabla As Recordset
Public tabla2 As Recordset
Public tabla3 As Recordset
Public AppPath As String
Public CDRom As String
Public CDSerial As Long
Public CDLabel As String
Public CDID As Long

Private Sub Main()

On Error GoTo Hell

If App.PrevInstance Then End

AppPath = Replace(App.Path & "\", "\\", "\")
Set bd = OpenDatabase(AppPath & "CDs.mdb")
frmMain.Show

Exit Sub
Hell:
    Select Case Err.Number
        Case 3024
            MsgBox "No se ha podido abrir la base de datos 'CDs.mdb'", vbCritical, "Error"
            End
    End Select
    Resume Next
End Sub

'//funcion que obtiene un numero unico para cada cd
Public Function SerialNumber() As Long
On Error Resume Next

Dim s As String
Dim ret As Long

s = String(30, Chr(0))

ret = mciSendString("info cdaudio identity", s, Len(s), 0)

SerialNumber = CLng(Trim(StripNulls(s)))

End Function

'//funcion que quita el caracter nulo
Function StripNulls(startStrg) As String
On Error Resume Next
Dim c As Integer, item As String
 c = 1
 Do
 If Mid(startStrg, c, 1) = Chr(0) Then
    item = Mid(startStrg, 1, c - 1)
    startStrg = Mid(startStrg, c + 1, Len(startStrg))
    StripNulls = item
    Exit Function
 End If
 c = c + 1
 Loop
End Function

Public Function SendMCIString(cmd As String) As Boolean
Static rc As Long
Static errStr As String * 200
rc = mciSendString(cmd, 0, 0, 0)
SendMCIString = (rc = 0)
End Function

Public Function DimeElNumeroDeSerie(ByVal Unidad As String, ByRef Etiqueta As String) As Long
On Error Resume Next

  Dim lVSN As Long, n As Long, s1 As String, s2 As String
    
    s1 = String$(255, Chr$(0))
    s2 = String$(255, Chr$(0))
    Unidad = Unidad & Chr$(0)
    Call GetVolumeInformation(Unidad, s1, 255, lVSN, 0, 0, s2, 255)
    Etiqueta = RTrim(StripNulls(s1))
  DimeElNumeroDeSerie = lVSN

End Function

Public Function EsUnaUnidadDeCDROM(Unidad As String) As Boolean
On Error Resume Next

  Dim lDrive As Long
  Dim szRoot As String
    
    szRoot = Unidad
    lDrive = GetDriveType(szRoot)
    If lDrive = DRIVE_CDROM Then
        EsUnaUnidadDeCDROM = True
    End If
    
End Function

Public Function GetDrives() As String
On Error Resume Next
Dim r As Long, allDrives As String, DriveType As Long
 allDrives = Space(64)
 r = GetLogicalDriveStrings(Len(allDrives), allDrives)
 GetDrives = Left(allDrives, r)
End Function

'//funcion que retrasa X tiempo el programa
Public Sub Espera(Segundos As Single)
On Error Resume Next
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub

Public Function DiskSize(ByVal Path As String) As Currency
On Error Resume Next
Dim BytesFreeToCaller As Currency, TotalBytes As Currency, TotalFreeBytes As Currency
Call GetDiskFreeSpaceEx(Path, BytesFreeToCaller, TotalBytes, TotalFreeBytes)
DiskSize = TotalBytes * 10000
End Function

'//funcion que devuelve el directorio anterior a uno dado
Public Function DirAnterior(ByVal Path As String) As String
    On Error Resume Next
    
    Do
        Path = Mid(Path, 1, Len(Path) - 1)
    Loop Until Right(Path, 1) = "\" Or Len(Path) < 3
    
    DirAnterior = Path
    
End Function

'//funcion que devuelve el directorio anterior a uno dado
Public Function DirActual(ByVal Path As String) As String
On Error Resume Next
    Dim i As Long
    Dim fin As Long
    
    i = 1
    While i <= Len(Path)
        If Mid(Path, i, 1) = "\" Then fin = i
        i = i + 1
    Wend
    
    DirActual = Mid(Path, fin + 1)
    
End Function

Public Function GetTempDir() As String
On Error Resume Next
Dim aux As String
Dim r As Long
aux = Space(255)
r = GetTempPath(255, aux)
aux = Mid(aux, 1, r)
If Right(aux, 1) <> "\" Then aux = aux & "\"
GetTempDir = aux
End Function

Public Function Filtra(ByVal Cadena As String) As String
Cadena = Replace(Cadena, "'", "")
Cadena = Replace(Cadena, "--", "")
Cadena = Replace(Cadena, ";", "")
Filtra = Cadena
End Function

Public Sub Ejecuta(Archivo As String, Optional Parametros As String)
Call ShellExecute(0, "open", Archivo, Parametros, App.Path, 1)
End Sub

