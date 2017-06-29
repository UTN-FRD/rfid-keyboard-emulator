Attribute VB_Name = "Module1"
Public Type TagIds
      TagType As Byte
      AntNum  As Byte
      Ids(11) As Byte
      End Type
  
Public Declare Function CommOpen Lib "Mr915ApiV10.dll" (ByRef hCom As Long, ByVal com_port As String) As Integer
Public Declare Function CommClose Lib "Mr915ApiV10.dll" (ByVal hCom As Long) As Integer
Public Declare Function SetBaudRate Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal BaudRate As Integer, ByVal NetAddr As Byte) As Integer

Public Declare Function ResetReader Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal NetAddr As Byte) As Integer
Public Declare Function GetFirmwareVersion Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef major As Byte, ByRef minor As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function SetRf Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal power As Byte, ByVal freq_type As Byte, ByVal NetAddr As Byte) As Integer
Public Declare Function GetRf Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef power As Byte, ByRef freq_type As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function SetAnt Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal ant As Byte, ByVal NetAddr As Byte) As Integer
Public Declare Function GetAnt Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef ant As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function IsoMultiTagIdentify Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef Count As Long, ByRef Value As TagIds, ByVal NetAddr As Byte) As Integer
Public Declare Function IsoMultiTagRead Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Long, ByRef Count As Long, ByRef Value As TagIds, ByVal NetAddr As Byte) As Integer
Public Declare Function IsoWriteTag Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Byte, ByVal Value As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function IsoReadWithID Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef TagID As Long, ByVal iAddr As Long, ByRef AntNum As Long, ByRef Value As Long, ByVal NetAddr As Byte) As Integer
Public Declare Function IsoWriteWithID Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef TagID As Long, ByVal iAddr As Long, ByVal Value As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function IsoLockTag Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Byte, ByVal NetAddr As Byte) As Integer
Public Declare Function IsoQueryLock Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Byte, ByRef status As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function IsoBlockWrite Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Byte, ByVal Length As Byte, ByRef Value As Long, ByVal NetAddr As Byte) As Integer
Public Declare Function IsoSigleTagRead Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Long, ByRef Value As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function ClearIDBuffer Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal iAddr As Byte) As Integer
Public Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lsString As String, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function Gen2MultiTagIdentify Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByRef Count As Long, ByRef Value As TagIds, ByVal NetAddr As Byte) As Integer

Public Declare Function Gen2WriteEPC Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal WordPtr As Byte, ByRef Value As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function Gen2LockTag Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal MemBank As Byte, ByVal NetAddr As Byte) As Integer
Public Declare Function Gen2KillTag Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal PassWord As Long, ByVal NetAddr As Byte) As Integer
Public Declare Function Gen2InitEPC Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal WordCount As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function Gen2Read Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal MemBank As Byte, ByVal WordPtr As Byte, ByVal wordcnt As Byte, ByRef Value As Byte, ByVal NetAddr As Byte) As Integer

Public Declare Function Gen2Write Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal MemBank As Byte, ByVal WordPtr As Byte, ByVal Value As Long, ByVal NetAddr As Byte) As Integer

Public Declare Function Gen2BlockWrite Lib "Mr915ApiV10.dll" (ByVal hCom As Long, ByVal MemBank As Byte, ByVal WordPtr As Byte, ByVal wordcnt As Byte, ByRef Value As Byte, ByVal NetAddr As Byte) As Integer

Public flagTag As Long
Public bButton2 As Boolean
Public hCom As Long
Public Timeflag  As Long
Public TimeCount As Integer


