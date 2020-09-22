VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "bcZipper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private WithEvents m_cUnzip As cUnzip
Attribute m_cUnzip.VB_VarHelpID = -1
Private WithEvents m_cZ As cZip
Attribute m_cZ.VB_VarHelpID = -1

Public Function ZipaFile(sFIle As String, sFinish As String) As Boolean
Dim Iret
On Error GoTo EH
'check to make sure the file and dir exist
Set m_cZ = New cZip

Iret = Dir(sFIle)
If Iret = "" Then
Err.Raise 53 ' file not found
End If

With m_cZ
.ZipFile = sFinish
.StoreFolderNames = False
.RecurseSubDirs = False
.ClearFileSpecs
.AddFileSpec sFIle
.Zip
End With

Set m_cZ = Nothing
ZipaFile = True
Exit Function
EH:
Set m_cZ = Nothing
ZipaFile = False
Exit Function
End Function

Public Function ZipaFolderEX(sDir As String, sFinish As String) As Boolean
Dim Iret
On Error GoTo EH
Set m_cZ = New cZip

Iret = Dir(sDir, vbDirectory)

If Iret = "" Then
Err.Raise 53 ' file not found
ZipaFolderEX = False
Exit Function
End If
If Right(sDir, 1) <> "\" Then sDir = sDir & "\"
With m_cZ
      .ZipFile = sFinish
      .BasePath = sDir
      .ClearFileSpecs
      .AddFileSpec "*.bmp"
      .AddFileSpec "*.jpg"
      .AddFileSpec "*.dat"
      .StoreFolderNames = True
      .RecurseSubDirs = True
      .StoreDirectories = True
      .Zip
End With

Set m_cZ = Nothing
ZipaFolderEX = True
Exit Function
EH:
Set m_cZ = Nothing
ZipaFolderEX = False
Exit Function
End Function

Public Function ZipaFolder(sDir As String, sFinish As String) As Boolean
Dim Iret
On Error GoTo EH
Set m_cZ = New cZip

Iret = Dir(sDir, vbDirectory)

If Iret = "" Then
Err.Raise 53 ' file not found
ZipaFolder = False
Exit Function
End If

With m_cZ
.ClearFileSpecs
.ZipFile = sFinish
.StoreFolderNames = False
.StoreDirectories = False
.RecurseSubDirs = False
.AddFileSpec sDir & "*.*"
.Zip
End With

Set m_cZ = Nothing
ZipaFolder = True
Exit Function
EH:
Set m_cZ = Nothing
ZipaFolder = False
Exit Function
End Function

Public Function UnzipaFileEX(sFIle As String, sDir As String) As Boolean
Dim i As Integer
Dim Iret
On Error GoTo EH
'check to make sure the file and dir exist
Set m_cUnzip = New cUnzip

Iret = Dir(sFIle)
If Iret = "" Then
Err.Raise 53 ' file not found
End If

Iret = Dir(sDir, vbDirectory)
If Iret = "" Then
Err.Raise 53 'file not found
End If

' Get the internal zip file directory:
m_cUnzip.ZipFile = sFIle
   
'select all files in the zip file
For i = 1 To m_cUnzip.FileCount
m_cUnzip.FileSelected(i) = True
Next i

'sfile = the name of the current file
'extract to the desired directory
With m_cUnzip
.UnzipFolder = sDir
.Unzip
End With
UnzipaFileEX = True

Exit Function
EH:
UnzipaFileEX = False
Exit Function
End Function

Public Function UnzipaFile(sFIle As String, sDir As String) As Boolean
Dim i As Integer
Dim Iret
On Error GoTo EH
'check to make sure the file and dir exist
Set m_cUnzip = New cUnzip

Iret = Dir(sFIle)
If Iret = "" Then
Err.Raise 53 ' file not found
End If

Iret = Dir(sDir, vbDirectory)
If Iret = "" Then
Err.Raise 53 'file not found
End If
m_cUnzip.ZipFile = sFIle
m_cUnzip.UnzipFolder = sDir
m_cUnzip.Directory
'select all files in the zip file
For i = 1 To m_cUnzip.FileCount
m_cUnzip.FileSelected(i) = True
m_cUnzip.UseFolderNames = True
Next i
'sfile = the name of the current file
'extract to the desired directory
m_cUnzip.Unzip
DoEvents
UnzipaFile = True

Exit Function
EH:
UnzipaFile = False
Exit Function
End Function

