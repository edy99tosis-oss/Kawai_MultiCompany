Attribute VB_Name = "mdlZip"
Option Explicit

Private Type tm_struct
    tm_seconds As Long
    tm_minutes As Long
    tm_hours   As Long
    tm_days    As Long
    tm_months  As Long
    tm_years   As Long
End Type

Private Type zip_fileinfo
    tmz         As tm_struct
    dosDate     As Long
    internal_fa As Long
    external_fa As Long
End Type

Private Declare Function zipOpen Lib "zlib.dll" (ByVal pathname As String, ByVal append As Long) As Long
Private Declare Function zipOpenNewFileInZip Lib "zlib.dll" (ByVal file As Long, ByVal filename As String, zipfi As zip_fileinfo, extrafield_local As Any, ByVal size_extrafield_local As Long, extrafield_global As Any, ByVal size_extrafield_global As Long, ByVal comment As String, ByVal method As Long, ByVal level As Long) As Long
Private Declare Function zipWriteInFileInZip Lib "zlib.dll" (ByVal file As Long, buf As Any, ByVal nLen As Long) As Long
Private Declare Function zipCloseFileInZip Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function zipClose Lib "zlib.dll" (ByVal file As Long, ByVal comment As String) As Long

Private Const Z_DEFLATED = 8
Private Const Z_BEST_COMPRESSION = 9

Private Type unz_file_info
    version            As Long
    version_needed     As Long
    Flag               As Long
    compression_method As Long
    dosDate            As Long
    crc                As Long
    compressed_size    As Long
    uncompressed_size  As Long
    size_filename      As Long
    size_file_extra    As Long
    size_file_comment  As Long
    disk_num_start     As Long
    internal_fa        As Long
    external_fa        As Long
    tmu_date           As tm_struct
End Type

Private Declare Function unzOpen Lib "zlib.dll" (ByVal path As String) As Long
Private Declare Function unzClose Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function unzGoToFirstFile Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function unzGoToNextFile Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function unzGetCurrentFileInfo Lib "zlib.dll" (ByVal file As Long, pfile_info As unz_file_info, ByVal szFileName As String, ByVal fileNameBufferSize As Long, extraField As Any, ByVal extraFieldBufferSize As Long, ByVal szComment As String, ByVal commentBufferSize As Long) As Long
Private Declare Function unzOpenCurrentFile Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function unzCloseCurrentFile Lib "zlib.dll" (ByVal file As Long) As Long
Private Declare Function unzReadCurrentFile Lib "zlib.dll" (ByVal file As Long, lpvoid As Any, ByVal nLen As Long) As Long

Const UNZ_OK = 0

Private Type FILETIME
    dwLowDateTime  As Long
    dwHighDateTime As Long
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As Any, lpLastAccessTime As Any, lpLastWriteTime As Any) As Long
Private Declare Function FileTimeToDosDateTime Lib "kernel32" (lpFileTime As FILETIME, lpFatDate As Any, lpFatTime As Any) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long

Private Const GENERIC_READ = &H80000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const GENERIC_WRITE = &H40000000
Private Const CREATE_ALWAYS = 2

Public Function CreateZipFile(psPath As String, cFiles As collection, psFiles As String, psFileName As String) As Boolean
   On Local Error GoTo err_handler
   Screen.MousePointer = vbHourglass
    
   Dim lngZip As Long
   Dim vItem As Variant
   Dim bArray() As Byte
   Dim hFile As Long
   Dim dwLen As Long
   Dim dwTemp As Long
   Dim zipfi As zip_fileinfo
   Dim ft1 As FILETIME
   Dim ft2 As FILETIME
   
   lngZip = zipOpen(psFileName, 0)
   If lngZip = 0 Then
      Exit Function
   End If

   For Each vItem In cFiles
      If Dir(psPath & vItem) = psFiles Then
         dwLen = FileLen(psPath & vItem)
         If dwLen <> 0 Then
            hFile = CreateFile(psPath & vItem, GENERIC_READ, FILE_SHARE_READ, ByVal 0, OPEN_EXISTING, 0, 0)
            ReDim bArray(dwLen)
            ReadFile hFile, bArray(0), dwLen, dwTemp, ByVal 0
            CloseHandle hFile

            GetFileTime hFile, ft1, ByVal 0, ByVal 0
            FileTimeToLocalFileTime ft1, ft2
            FileTimeToDosDateTime ft2, ByVal (VarPtr(zipfi.dosDate) + 2), ByVal VarPtr(zipfi.dosDate)

            zipOpenNewFileInZip lngZip, vItem, zipfi, ByVal 0, 0, ByVal 0, 0, vbNullString, Z_DEFLATED, Z_BEST_COMPRESSION
            zipWriteInFileInZip lngZip, bArray(0), dwLen
            zipCloseFileInZip lngZip
         End If
      End If
   Next

   zipClose lngZip, vbNullString
   CreateZipFile = True
    
err_exit:
   Screen.MousePointer = vbDefault
   'zipClose lngZip, vbNullString
   Exit Function
err_handler:
   MsgBox "Err. Number : " & Err.number & "  " & vbNewLine & "Err. Description : " & Err.Description & "", vbCritical, "Change Apply Setting Failed"
   Err.clear
   Resume err_exit
End Function
