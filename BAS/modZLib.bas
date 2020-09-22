Attribute VB_Name = "modzLib"
'===========================================================================
'
' modZLib
'
'---------------------------------------------------------------------------
'
' Name:           modZLib
' Author:         Daniel Keep
' Contact:        shortcircuitfky@hotmail.com
'                 When you contact me, please put [VB] in the subject line
'                 otherwise I might not read it!
' Version:        1.0
' Created:        2003-05-18
' Description:    Wrapper for zlib.dll
' Requirements:   zlib.dll version 1.1.4 or higher (may work with earlier
'                 versions, but I can't verify this)
' Credits:        Constants and declarations adapted from zlib.h
' Changes    1.0: Initial version.  Implements compress(), compress2(),
'                 uncompress(), and for the first time (that I've seen in
'                 a VB wrapper for zlib) adler32() and crc32().  All methods
'                 have versions for both byte arrays and strings, as well
'                 as simple and 'Ex'tended versions for more control.
'                 Enjoy.
'
' License:
'   You are given non-exclusive permission to use and modify this source
'   code provided that  credit is given  to the original  author(s), and
'   that if this code is used for commercial purposes that the author is
'   informed of this  use.  It is requested, although not  required that
'   the author(s)  be informed  of the use of  this source  code for any
'   purposes.
'   This  source code  is provided AS  IS, and  comes with no  warranty,
'   implied or otherwise.  You use it at your own risk.
'
'                       Copyright © 2003 Daniel Keep
'        zlib Copyright © 1995-2002 Jean-loup Gailly and Mark Adler
'===========================================================================
Option Explicit

'===========================================================================
'===========================================================================
'===========================================================================

' Win32 Declares
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

' zlib Declares

' If you want these declare statements public to your project, make sure
' you put "ZLIB_PUBLIC" into your conditional compilation arguments list.
#If ZLIB_PUBLIC Then

'int compress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Public Declare _
  Function zlibCompress _
  Lib "zlib.dll" _
  Alias "compress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long

' int compress2 (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen, int level);

Public Declare _
  Function zlibCompress2 _
  Lib "zlib.dll" _
  Alias "compress2" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long, _
    ByVal Level As Long _
  ) As Long

'int uncompress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Public Declare _
  Function zlibUncompress _
  Lib "zlib.dll" _
  Alias "uncompress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long

' uLong adler32 (uLong adler, const Bytef *buf, uInt len);

Public Declare _
  Function zlibAdler32 _
  Lib "zlib.dll" _
  Alias "adler32" _
  ( _
    ByVal adler As Long, _
    ByRef buf As Any, _
    ByVal Length As Long _
  ) As Long


' uLong crc32 (uLong crc, const Bytef *buf, uInt len);

Public Declare _
  Function zlibCRC32 _
  Lib "zlib.dll" _
  Alias "crc32" _
  ( _
    ByVal crc As Long, _
    ByRef buf As Any, _
    ByVal Length As Long _
  ) As Long

#Else

'int compress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Private Declare _
  Function zlibCompress _
  Lib "zlib.dll" _
  Alias "compress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long

' int compress2 (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen, int level);

Private Declare _
  Function zlibCompress2 _
  Lib "zlib.dll" _
  Alias "compress2" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long, _
    ByVal Level As Long _
  ) As Long

'int uncompress (Bytef *dest, uLongf *destLen, const Bytef *source, uLong sourceLen);

Private Declare _
  Function zlibUncompress _
  Lib "zlib.dll" _
  Alias "uncompress" _
  ( _
    ByRef dest As Any, _
    ByRef destLen As Long, _
    ByRef Source As Any, _
    ByVal sourceLen As Long _
  ) As Long

' uLong adler32 (uLong adler, const Bytef *buf, uInt len);

Private Declare _
  Function zlibAdler32 _
  Lib "zlib.dll" _
  Alias "adler32" _
  ( _
    ByVal adler As Long, _
    ByRef buf As Any, _
    ByVal Length As Long _
  ) As Long


' uLong crc32 (uLong crc, const Bytef *buf, uInt len);

Private Declare _
  Function zlibCRC32 _
  Lib "zlib.dll" _
  Alias "crc32" _
  ( _
    ByVal crc As Long, _
    ByRef buf As Any, _
    ByVal Length As Long _
  ) As Long

#End If

'===========================================================================
'===========================================================================
'===========================================================================
' Private Constants

' As with the declares, if you want these constants to be public, make sure
' you add ZLIB_PUBLIC to your conditional compilation arguments

#If ZLIB_PUBLIC Then

' Return codes for compression/decompression functions.  Negative values are
' errors, positive values are used for special but normal events.
Public Const Z_OK As Long = 0
Public Const Z_STREAM_END As Long = 1
Public Const Z_NEED_DICT As Long = 2
Public Const Z_ERRNO As Long = -1
Public Const Z_STREAM_ERROR As Long = -2
Public Const Z_DATA_ERROR As Long = -3
Public Const Z_MEM_ERROR As Long = -4
Public Const Z_BUF_ERROR As Long = -5
Public Const Z_VERSION_ERROR As Long = -6

' Compression levels
Public Const Z_NO_COMPRESSION As Long = 0
Public Const Z_BEST_SPEED As Long = 1
Public Const Z_BEST_COMPRESSION As Long = 9
Public Const Z_DEFAULT_COMPRESSION As Long = -1

#Else

' Return codes for compression/decompression functions.  Negative values are
' errors, positive values are used for special but normal events.
Private Const Z_OK As Long = 0
Private Const Z_STREAM_END As Long = 1
Private Const Z_NEED_DICT As Long = 2
Private Const Z_ERRNO As Long = -1
Private Const Z_STREAM_ERROR As Long = -2
Private Const Z_DATA_ERROR As Long = -3
Private Const Z_MEM_ERROR As Long = -4
Private Const Z_BUF_ERROR As Long = -5
Private Const Z_VERSION_ERROR As Long = -6

' Compression levels
Private Const Z_NO_COMPRESSION As Long = 0
Private Const Z_BEST_SPEED As Long = 1
Private Const Z_BEST_COMPRESSION As Long = 9
Private Const Z_DEFAULT_COMPRESSION As Long = -1

#End If

' Default compression level
Private Const DEFAULT_LEVEL As Long = Z_DEFAULT_COMPRESSION

'===========================================================================
'===========================================================================
'===========================================================================
' Public Methods

'===========================================================================
'===========================================================================
' zlibCompressBytes()
'  Compresses the specified byte array

Public Function zlibCompressBytes( _
  ByRef Data() As Byte, _
  Optional ByRef Level As Long = DEFAULT_LEVEL _
) As Byte()
  
  ' Make a temporary buffer for the compressed output
  Dim cTemp() As Byte
  Dim cTempLen As Long
  
  ' Find the length of the data
  Dim nDataLen As Long
  nDataLen = 0&
  
  On Error Resume Next
  nDataLen = UBound(Data) + 1&
  On Error GoTo 0
  
  ' Now compress it
  zlibCompressBytesEx Data, nDataLen, cTemp, cTempLen, Level
  
  ' Finally, return the array
  zlibCompressBytes = cTemp
  
End Function

'===========================================================================
'===========================================================================
' zlibCompressBytesEx()
'  Compresses the specified byte array

Public Function zlibCompressBytesEx( _
  ByRef Data() As Byte, _
  ByRef DataLen As Long, _
  ByRef Compressed() As Byte, _
  ByRef CompressedLen As Long, _
  ByRef Level As Long _
) As Long
  
  Dim l_return As Long
  
  #If DEBUG_ENABLED And DEBUG_MID Then
    On Error GoTo zlibCompressBytesEx_Error
    dbgEnter "modZLib.zlibCompressBytesEx()"
  #End If
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
  
  '-------------------------------------------------------------------------
  ' Make a buffer for the compressed data
  Dim cDest() As Byte, cDestLen&
  cDestLen = CLng((DataLen + 12&) * 1.1)
  ReDim cDest(cDestLen - 1&)
  
  ' Return value
  Dim nReturn&
  
  '-------------------------------------------------------------------------
  ' Ok, call zlibCompress()
  nReturn = zlibCompress(cDest(0&), cDestLen, Data(0&), DataLen)
  
  '-------------------------------------------------------------------------
  ' Did we succeed?
  If nReturn <> Z_OK Then
    
    ' Nope.
    l_return = nReturn
    GoTo Terminate
    
  End If
  
  ' If we got here, we succeeded.  Return the compressed data & length
  ReDim Compressed(cDestLen - 1&)
  CopyMemory Compressed(0&), cDest(0&), cDestLen
  CompressedLen = cDestLen
  
  l_return = Z_OK
  
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Terminate:
  #If DEBUG_ENABLED And DEBUG_MID Then
    On Error Resume Next
    dbgLeave
  #End If
  zlibCompressBytesEx = l_return
  Exit Function

#If DEBUG_ENABLED And DEBUG_MID Then
zlibCompressBytesEx_Error:
    dbgPrintVBError Err
    dbgDumpStack
    dbgDumpMachineInfo
    GoTo Terminate
#End If

End Function

'===========================================================================
'===========================================================================
' zlibCompressString()
'  Compresses the specified string

Public Function zlibCompressString( _
  ByRef Data As String, _
  Optional ByRef Level As Long = DEFAULT_LEVEL _
) As String
  
  ' Convert the string to a byte array
  Dim cSource() As Byte
  Dim cDest() As Byte
  
  cSource = StrConv(Data, vbFromUnicode)
  cDest = zlibCompressBytes(cSource, Level)
  zlibCompressString = StrConv(cDest, vbUnicode)
  
End Function

'===========================================================================
'===========================================================================
' zlibCompressStringEx()
'  Compresses the specified string

Public Function zlibCompressStringEx( _
  ByRef Data As String, _
  ByRef Compressed As String, _
  ByRef Level As Long _
) As Long
  
  Dim cSource() As Byte, cSourceLen As Long
  Dim cDest() As Byte, cDestLen As Long
  Dim nResult As Long
  
  cSource = StrConv(Data, vbFromUnicode)
  cSourceLen = Len(Data)
  
  nResult = zlibCompressBytesEx(cSource, cSourceLen, cDest, cDestLen, Level)
  
  Compressed = StrConv(cDest, vbUnicode)
  zlibCompressStringEx = nResult
  
End Function

'===========================================================================
'===========================================================================
' zlibUncompressBytes()
'  Decompresses the specified byte array

Public Function zlibUncompressBytes( _
  ByRef Data() As Byte, _
  ByRef UncompressedLen As Long _
) As Byte()
  
  Dim cDest() As Byte
  Dim nDataLen As Long
  
  nDataLen = 0&
  On Error Resume Next
  nDataLen = UBound(Data) + 1&
  On Error GoTo 0
  
  zlibUncompressBytesEx Data, nDataLen, cDest, UncompressedLen
  
  zlibUncompressBytes = cDest
  
End Function

'===========================================================================
'===========================================================================
' zlibUncompressBytesEx()
'  Decompresses the specified byte array

Public Function zlibUncompressBytesEx( _
  ByRef Data() As Byte, _
  ByRef DataLen As Long, _
  ByRef Uncompressed() As Byte, _
  ByRef UncompressedLen As Long _
) As Long

  Dim l_return As Long
  
  #If DEBUG_ENABLED And DEBUG_MID Then
    On Error GoTo zlibUncompressBytesEx_Error
    dbgEnter "modZLib.zlibUncompressBytesEx()"
  #End If
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
  
  '-------------------------------------------------------------------------
  ' Define some variables
  Dim cDest() As Byte, cDestLen As Long
  Dim nReturn As Long
  
  cDestLen = UncompressedLen
  ReDim cDest(cDestLen)
  
  '-------------------------------------------------------------------------
  ' Make the call to zlib
  nReturn = zlibUncompress(cDest(0&), cDestLen, Data(0&), DataLen)
  
  '-------------------------------------------------------------------------
  ' Did we succeed?
  If nReturn <> Z_OK Then
    
    l_return = nReturn
    GoTo Terminate
    
  End If
  
  '-------------------------------------------------------------------------
  ' Ok: return the uncompressed data
  ReDim Uncompressed(cDestLen - 1&)
  CopyMemory Uncompressed(0&), cDest(0&), cDestLen
  UncompressedLen = cDestLen
  
  l_return = nReturn
  
'---------------------------------------------------------------------------
'---------------------------------------------------------------------------
Terminate:
  #If DEBUG_ENABLED And DEBUG_MID Then
    On Error Resume Next
    dbgLeave
  #End If
  zlibUncompressBytesEx = l_return
  Exit Function

#If DEBUG_ENABLED And DEBUG_MID Then
zlibUncompressBytesEx_Error:
    dbgPrintVBError Err
    dbgDumpStack
    dbgDumpMachineInfo
    GoTo Terminate
#End If

End Function

'===========================================================================
'===========================================================================
' zlibUncompressString()
'  Decompresses the specified string

Public Function zlibUncompressString( _
  ByRef Data As String, _
  ByRef UncompressedLen As Long _
) As String
  
  Dim cData() As Byte
  Dim cDest() As Byte
  
  cData = StrConv(Data, vbFromUnicode)
  cDest = zlibUncompressBytes(cData, UncompressedLen)
  zlibUncompressString = StrConv(cDest, vbUnicode)
  
End Function

'===========================================================================
'===========================================================================
' zlibUncompressStringEx()
'  Decompresses the specified string

Public Function zlibUncompressStringEx( _
  ByRef Data As String, _
  ByRef Uncompressed As String, _
  ByRef UncompressedLen As Long _
) As Long
  
  Dim cData() As Byte, cDataLen As Long
  Dim cDest() As Byte
  Dim nReturn As Long
  
  cData = StrConv(Data, vbFromUnicode)
  cDataLen = Len(Data)
  
  nReturn = zlibUncompressBytesEx(cData, cDataLen, cDest, UncompressedLen)
  
  Uncompressed = StrConv(cDest, vbUnicode)
  zlibUncompressStringEx = nReturn
  
End Function

'===========================================================================
'===========================================================================
' zlibAdler32Bytes()
'  Calculates the Adler32 checksum of the specified byte array

Public Function zlibAdler32Bytes( _
  ByRef Data() As Byte _
) As Long

  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nDataLen As Long
  Dim nChecksum As Long
  
  nDataLen = 0&
  On Error Resume Next
  nDataLen = UBound(Data) + 1&
  On Error Resume Next
  
  nChecksum = zlibAdler32(0&, ByVal 0&, 0&)
  nChecksum = zlibAdler32(nChecksum, Data(0&), nDataLen)
  
  zlibAdler32Bytes = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibAdler32BytesEx()
'  Calculates the progressive Adler32 checksum of the specified byte array

Public Function zlibAdler32BytesEx( _
  ByRef Data() As Byte, _
  ByRef DataLen As Long, _
  ByRef Adler32 As Long _
) As Long
  
  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nChecksum As Long
  
  nChecksum = zlibAdler32(Adler32, Data(0&), DataLen)
  
  zlibAdler32BytesEx = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibAdler32String()
'  Calculates the Adler32 checksum of the specified string

Public Function zlibAdler32String( _
  ByRef Data As String _
) As Long
  
  Dim cData() As Byte
  Dim nReturn As Long
  
  cData = StrConv(Data, vbFromUnicode)
  nReturn = zlibAdler32Bytes(cData)
  zlibAdler32String = nReturn
  
End Function

'===========================================================================
'===========================================================================
' zlibAdler32StringEx()
'  Calculates the progressive Adler32 checksum of the specified string

Public Function zlibAdler32StringEx( _
  ByRef Data As String, _
  ByRef Adler32 As Long _
) As Long
  
  Dim cData() As Byte, cDataLen As Long
  Dim nReturn As Long
  
  cData = StrConv(Data, vbFromUnicode)
  cDataLen = Len(Data)
  nReturn = zlibAdler32BytesEx(cData, cDataLen, Adler32)
  zlibAdler32StringEx = nReturn
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32Bytes()
'  Calculates the CRC32 checksum of the specified byte array

Public Function zlibCRC32Bytes( _
  ByRef Data() As Byte _
) As Long

  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nDataLen As Long
  Dim nChecksum As Long
  
  nDataLen = 0&
  On Error Resume Next
  nDataLen = UBound(Data) + 1&
  On Error Resume Next
  
  nChecksum = zlibCRC32(0&, ByVal 0&, 0&)
  nChecksum = zlibCRC32(nChecksum, Data(0&), nDataLen)
  
  zlibCRC32Bytes = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32BytesEx()
'  Calculates the progressive CRC32 checksum of the specified byte array

Public Function zlibCRC32BytesEx( _
  ByRef Data() As Byte, _
  ByRef DataLen As Long, _
  ByRef CRC32 As Long _
) As Long
  
  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nChecksum As Long
  
  nChecksum = zlibCRC32(CRC32, Data(0&), DataLen)
  
  zlibCRC32BytesEx = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32Mem()
'  Calculates the CRC32 checksum on a block of memory

Public Function zlibCRC32Mem( _
  ByRef Source As Long, _
  ByRef Length As Long _
) As Long

  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nChecksum As Long
  
  nChecksum = zlibCRC32(0&, ByVal 0&, 0&)
  nChecksum = zlibCRC32(nChecksum, ByVal Source, Length)
  
  zlibCRC32Mem = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32MemEx()
'  Calculates the progressive CRC32 checksum on a block of memory

Public Function zlibCRC32MemEx( _
  ByRef Source As Long, _
  ByRef Length As Long, _
  ByRef CRC32 As Long _
) As Long

  '-------------------------------------------------------------------------
  ' Declare some vars
  Dim nChecksum As Long
  
  nChecksum = zlibCRC32(CRC32, ByVal Source, Length)
  
  zlibCRC32MemEx = nChecksum
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32String()
'  Calculates the CRC32 checksum of the specified string

Public Function zlibCRC32String( _
  ByRef Data As String _
) As Long
  
  Dim cData() As Byte
  Dim nReturn As Long
  
  cData = StrConv(Data, vbFromUnicode)
  nReturn = zlibCRC32Bytes(cData)
  zlibCRC32String = nReturn
  
End Function

'===========================================================================
'===========================================================================
' zlibCRC32StringEx()
'  Calculates the progressive CRC32 checksum of the specified string

Public Function zlibCRC32StringEx( _
  ByRef Data As String, _
  ByRef CRC32 As Long _
) As Long
  
  Dim cData() As Byte, cDataLen As Long
  Dim nReturn As Long
  
  cData = StrConv(Data, vbFromUnicode)
  cDataLen = Len(Data)
  nReturn = zlibCRC32BytesEx(cData, cDataLen, CRC32)
  zlibCRC32StringEx = nReturn
  
End Function



