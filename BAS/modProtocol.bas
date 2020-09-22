Attribute VB_Name = "modProtocol"
'// --------------------------------------------------------------------------
'// ISPN Byte Delimiting Demo
'// --------------------------------------------------------------------------
'// modProtocol
'// ***********
'// Converts protocol commands to actual protocol data, and also extracts data
'// fields from recieved protocol data.
'//
'// ©2003 Matthew Hall
'//       matthew@ispn-online.co.uk
'//       www.ispn-online.co.uk
'//
'// v1.0 MH
'//
'// 1.0: MH, Initial Version
'//
'// ToDo: Nothing.
'//
'// 2nd SEPTEMBER 2003
'// 16:09 GMT
'// --------------------------------------------------------------------------

'Force variable declaration
Option Explicit

'**ISPN Protocol Module**
'--Converts data to Protocol strings--

Public Function LogonProtocol_100(sLocalHandle As String, sLocalPassword As String, iServerPIN As Integer, sSoftwareVersion As String) As String
    
    LogonProtocol_100 = "100 " & ConvArgSet(ConvArg(sLocalHandle) & ConvArg(sLocalPassword) & ConvArg(iServerPIN) & ConvArg(sSoftwareVersion) & ConvArg(App.Major & "." & App.Minor & "." & App.Revision))

End Function

Public Function ExtractProtData(ByVal sPackData As String, ByRef sDestDynamicArry() As String, iExpectedArraySize As Integer) As Boolean

    On Error GoTo extractfail:
    
    '**Extract all fields from sPackData to an array**
    
    Dim sArgSet As String
    
    If sPackData = "" Then GoTo extractfail
    
    'See if argument set is compressed
    
    If Mid$(sPackData, 1, 1) = "1" Then
    
        '--Argument set is compressed--
        
        Dim sPlayString As String
        
        sPlayString = Mid$(sPackData, 2) 'Change our play string to not include the first char
                                         'which is the compression flag. We have already processed
                                         'so we just chop it off so it doesnt get in the way.
        
        'The next field is Packet Length in Bytes.
        'We dont need this, its used in the TCP buffer
        'only, for delimiting packets. So, we need to
        'look for the trailing space after this field
        'in order to advance to the next field.
                                
        Do
            
            
            If Mid$(sPlayString, 1, 1) = " " Then
                
                sPlayString = Mid$(sPlayString, 2)
                Exit Do
                
            Else
                
                sPlayString = Mid$(sPlayString, 2)
                
            End If
        
        Loop
        
        'Now we are on the next field, which is
        'Inflated Data Size. We need to pass this
        'value to the ZLib decompressor.
        
        Dim sRetVal As String
        
        Do
            
            
            If Mid$(sPlayString, 1, 1) = " " Then
                
                
                sPlayString = Mid$(sPlayString, 2)
                Exit Do
                
            Else
                
                sRetVal = sRetVal + Mid$(sPlayString, 1, 1) 'Add character to return value
                sPlayString = Mid$(sPlayString, 2)
            
            End If
        
        Loop
        
        'Decompress protocol arguments using ZLib
        sPlayString = modzLib.zlibUncompressString(sPlayString, CLng(sRetVal))
         
        sArgSet = sPlayString
    
    Else
        
        '--Argument set is not compressed--
        
        sArgSet = Mid$(sPackData, 2) 'Change our string to not include the first char
                                     'which is the compression flag. We have already
                                     'processed so we just chop it off so it doesnt
                                     'get in the way.
        Do
                        
            If Mid$(sArgSet, 1, 1) = " " Then
                
                sArgSet = Mid$(sArgSet, 2)
                Exit Do
                
            Else
                
                sArgSet = Mid$(sArgSet, 2)
                
            End If
        
        Loop
        
    End If
    
    'We have an uncompressed, byte delimited argument set that we now need to split into
    'an array.
    
    'Make sure array is empty
    Erase sDestDynamicArry()
    
    Dim sArgLen As String
    Dim lngArgNum As Long
        
    sArgLen = ""
    lngArgNum = 0
    
    Do
        
        'Read byte length

        If Mid$(sArgSet, 1, 1) = " " Then
            
            sArgSet = Mid$(sArgSet, 2)
            
            'Read argument
            ReDim Preserve sDestDynamicArry(lngArgNum)
            
            sDestDynamicArry(lngArgNum) = Mid$(sArgSet, 1, CLng(sArgLen))
            
            lngArgNum = lngArgNum + 1
            
            'See if this is the last field..
            If Len(sArgSet) = CLng(sArgLen) Then
                
                'No more fields
                Exit Do
            
            Else
                
                'More fields to come so crop the field we have just
                'read into the array element off, so we can read the
                'next.
                sArgSet = Mid$(sArgSet, CLng(sArgLen) + 1)
                sArgLen = ""
                
            End If
            
        Else
            
            sArgLen = sArgLen + Mid$(sArgSet, 1, 1) 'Add character to return value
            sArgSet = Mid$(sArgSet, 2)
            
        End If
        
    Loop
    
    If IsArrayEmpty(sDestDynamicArry) = True Then GoTo extractfail

    If Not UBound(sDestDynamicArry) = iExpectedArraySize Then GoTo extractfail
    
    ExtractProtData = True
    
    Exit Function

extractfail:

    ExtractProtData = False
    
End Function

Public Function IsArrayEmpty(aTestArray() As String) As Boolean
    
    On Error GoTo arryisempty
    
    Dim iRetVal As Integer
    
    iRetVal = UBound(aTestArray)
    
    IsArrayEmpty = False
    
    Exit Function

arryisempty:
    
    IsArrayEmpty = True

End Function

Private Function ConvArg(vArg As Variant) As String
    
    '**Converts argument to byte delimited field**
    ConvArg = CStr(Len(vArg) & " " & vArg)
    
End Function

Private Function ConvArgSet(sArgumentSet As String) As String
        
    '**Converts argument set to byte delimited field**
    '**and compresses if ZLib offers decent compression**
            
    Dim sComp As String
    sComp = sArgumentSet
    
    sComp = modzLib.zlibCompressString(sComp, 9)
    
    'Compare compression\no compression sizes
    
    Dim sWithComp As String, sWithoutComp As String
    
    sWithComp = "1" & (Len(CStr(Len(sArgumentSet))) + Len(sComp) + 1) & " " & Len(sArgumentSet) & " " & sComp
    sWithoutComp = "0" & Len(sArgumentSet) & " " & sArgumentSet
    
    If Not Len(sWithComp) <= Len(sWithoutComp) Then
        
        'Compressed size is the same or greater than uncompressed
        'size, so dont return compressed.
        ConvArgSet = sWithoutComp
        
    Else
        
        'Compressed size is smaller than uncompressed size,
        'so return compressed argument set.
        ConvArgSet = sWithComp
    
    End If
    
End Function
