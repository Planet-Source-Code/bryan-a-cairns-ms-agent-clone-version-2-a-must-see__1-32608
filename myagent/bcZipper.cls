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

Public Function ZipaFile(sFile As String, sFinish As String) As Boolean
Dim Iret
On Error GoTo EH
'check to make sure the file and dir exist
Set m_cZ = New cZip

Iret = Dir(sFile)
If Iret = "" Then
Err.Raise 53 ' file not found
End If

With m_cZ
.ZipFile = sFinish
.StoreFolderNames = False
.RecurseSubDirs = False
.ClearFileSpecs
.AddFileSpec sFile
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
.ZipFile = sFinish
.StoreFolderNames = False
.RecurseSubDirs = False
.ClearFileSpecs
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

Public Function UnzipaFile(sFile As String, sDir As String) As Boolean
Dim I As Integer
Dim Iret
On Error GoTo EH
'check to make sure the file and dir exist
Set m_cUnzip = New cUnzip

Iret = Dir(sFile)
If Iret = "" Then
Err.Raise 53 ' file not found
End If

Iret = Dir(sDir, vbDirectory)
If Iret = "" Then
Err.Raise 53 'file not found
End If

' Get the internal zip file directory:
m_cUnzip.ZipFile = sFile
m_cUnzip.UseFolderNames = True
'select all files in the zip file
For I = 1 To m_cUnzip.FileCount
m_cUnzip.FileSelected(I) = True
Next I

'sfile = the name of the current file
'extract to the desired directory
m_cUnzip.UnzipFolder = sDir
m_cUnzip.Unzip
UnzipaFile = True

Exit Function
EH:
UnzipaFile = False
Exit Function
End Function

