VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Server"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSend 
      Caption         =   "< Send as packet"
      Height          =   1590
      Left            =   4110
      TabIndex        =   8
      Top             =   540
      Width           =   1920
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Text            =   "Text2"
      Top             =   1650
      Width           =   1485
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   2520
      TabIndex        =   6
      Text            =   "123"
      Top             =   1230
      Width           =   1485
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   750
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Left            =   2520
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   330
      Width           =   1485
   End
   Begin MSWinsockLib.Winsock wsServer 
      Left            =   6780
      Top             =   75
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   5678
   End
   Begin VB.CommandButton cmdKill 
      Caption         =   "Disconnect"
      Height          =   900
      Left            =   225
      TabIndex        =   2
      Top             =   1320
      Width           =   2040
   End
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   225
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2610
      Width           =   6840
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "Reset\Listen"
      Height          =   915
      Left            =   225
      TabIndex        =   0
      Top             =   330
      Width           =   2040
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Log:"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   2340
      Width           =   315
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'// --------------------------------------------------------------------------
'// ISPN Byte Delimiting Demo
'// --------------------------------------------------------------------------
'// frmServer
'// *********
'// The server side of the demo. Accepts a connection from a client.
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

'TCP Recieve Buffer
Dim sTCPBuffer As String

Private Sub SendWSData(vData As Variant, sContext As String)
    
    '**TCP Send buffer**
    On Error GoTo ws_err
    
    Dim vTCPSendBuffer As Variant
    
    'Only send TCP data in chunks of 4096 Bytes at any one time
    
    vTCPSendBuffer = vData
    
snd_data:
    
    If Len(vTCPSendBuffer) > 4096 Then
        
        DoEvents
        wsServer.SendData Mid(vTCPSendBuffer, 1, 4096)
        vTCPSendBuffer = Mid(vTCPSendBuffer, 4096)
        GoTo snd_data
        
    Else
        
        wsServer.SendData vTCPSendBuffer
    
    End If
    
    Exit Sub
    
ws_err:
    
    MsgBox Err & ":" & Err.Description & vbCrLf & sContext
    
End Sub

Private Sub cmdKill_Click()
    
    wsServer.Close
        
    txtOutput.Text = txtOutput.Text & "socket closed" & vbCrLf

End Sub

Private Sub cmdListen_Click()
    
    wsServer.Close
    wsServer.Listen
    
    txtOutput.Text = txtOutput.Text & "socket open" & vbCrLf

End Sub

Private Sub cmdSend_Click()
    
    SendWSData modProtocol.LogonProtocol_100(Text1, Text2, CInt(Text3), Text4), "test"

End Sub

Private Sub wsServer_Close()
    
    wsServer.Close
    
    txtOutput.Text = txtOutput.Text & "socket closed" & vbCrLf

End Sub

Private Sub wsServer_ConnectionRequest(ByVal requestID As Long)
    
    wsServer.Close

    wsServer.Accept requestID
    
    txtOutput.Text = txtOutput.Text & "socket connected" & vbCrLf

End Sub

Public Sub wsServer_DataArrival(ByVal bytesTotal As Long)

    '**Deal with recieved winsock data**
    Dim sPacketData As String
    
    'Get the packet data and assign to sPacketData
    wsServer.GetData sPacketData
    
    '****TCP Buffer****
tcpbuffer:
    

    If sTCPBuffer = "" Then
        
        'Buffer previously empty, we do not need to append
        'this data to the buffer. We still need to make sure
        'the data is not truncated, ie. packet is x Bytes long
        
        'If there is less than 6 characters in the packet,
        'we cant find the length of the packet, so buffer the data
        'till the next one comes through.
        '
        '6 Chars = PID[3] + SP1[1] + CFLAG[1] + PLEN[At Least 1 byte]
        '
        If Len(sPacketData) < 6 Then
            
            'Less than 6 chars, so buffer till next packet is recieved
            sTCPBuffer = sPacketData
            Exit Sub
            
        End If
        
    Else
        
        'The buffer was NOT previously empty, in other words, last time
        'we had data recieved, there wasnt enough data to process, so it
        'was buffered.
        '
        'We have now recieved the rest (or part of the rest) of the
        'data, so we can prefix the buffered data to the start of the
        'new data.
        
        'Prefix buffered data
        sPacketData = sTCPBuffer & sPacketData
        
        'Clear buffer
        sTCPBuffer = ""
        
        'Reprocess
        GoTo tcpbuffer
        
    End If
        
        
    'There is a pointer telling us how long the packet is
    'after the protocol ID. Eg..
    '
    ' PID
    '(Protocol Command ID)
    '  | SP1   SP2
    '  | |    /
    ' 100 030 ?????????????????????????????
    '     ||| <---Packet Data(30 Bytes)--->
    '     |\\           \_PDATA_/
    '    /  \\
    '   |    \\-> PLEN (Packet Length in Bytes)
    '   /
    '  |
    'CFLAG
    'Compression
    'Flag (1\0)
    '
    'We dont need to worry about the CFLAG compression flag or
    'PID protocol ID here, it is handled in modProtocol.ExtractProtocolData
    'and DealWithPacket(). So start reading from the next field (PLEN)
    
    Dim sPlayString As String
    Dim sRetVal As String
    
    sPlayString = Mid$(sPacketData, 6)
    sRetVal = ""
    
    Do
                
        If Mid$(sPlayString, 1, 1) = " " Then
            
            'We found the delimiting space, and therefore
            'the full field data will be stored in sRetVal.
            '
            'We set the playstring here to be the start of
            'the next field (PDATA)
            sPlayString = Mid$(sPlayString, 2)
            Exit Do
            
        Else
            
            sRetVal = sRetVal + Mid$(sPlayString, 1, 1) 'Add character to return value
            sPlayString = Mid$(sPlayString, 2) 'Trim the playstring
            
            If sPlayString = "" Then
                
                'Run out of string to read..
                'Buffer it for next time..
                            
                sTCPBuffer = sPacketData
                Exit Sub
                 
            End If
            
        End If
    
    Loop
     
    'We have found the PLEN packet length field (now stored in sRetVal)
    'so now it is possible to determine if we have recieved an entire
    'packet or not.
    
    Dim lngMinLen As Long
    
    lngMinLen = 6 + Len(sRetVal) + CLng(sRetVal)
        
    If Len(sPacketData) < lngMinLen Then
        
        'Data doesnt contain at least one full packet
        
        'Buffer data
        sTCPBuffer = sPacketData
        
        'Exit routine to wait for the rest of the data
        Exit Sub
        
    End If
        
    
    'Okay, so we have at least one full packet.
    '
    'Now we need to see if there is more than one full packet..
    
    Dim sThePacket As String
    
    If Len(sPacketData) > lngMinLen Then
        
        'Aha.. so there IS more than one packet.
        '
        'Get our complete packet and store it in
        'sThePacket, and clip the rest off into
        'sPacketData for reprocessing.
        
        sThePacket = Mid$(sPacketData, 1, lngMinLen)
        
        sPacketData = Mid$(sPacketData, lngMinLen + 1)
        
    Else
    
        'There is only the one complete packet..
        
        sThePacket = sPacketData
        
        sPacketData = ""
        
    End If
    
    
    'Process the complete packet in sThePacket, and if there
    'was more than one packet, reprocess it.
         
processpacket:
    
    Call ShowIN(sThePacket)
    
    'Extract the fields from the protocol
    Dim strFields() As String
    
    'NOTE: THE EXPECTED ARRAY SIZE WILL CHANGE ACCORDING TO THE AMOUNT
    'OF DATA FIELDS THAT YOU SEND. IN THIS EXAMPLE THERE WILL ALWAYS BE 5 (Protocol 100 has 5 fields),
    'WHICH IN ARRAY TERMS IS 4 (BASE ZERO, 0,1,2,3,4 = 5 ELEMENTS). YOU COULD
    'USE A SELECT CASE ON THE PROTOCOL ID FIELD (SEE DOCUMENTATION) TO
    'DETERMINE THE NUMBER OF DATA FIELDS WE ARE EXPECTING.
    
    'ALSO NOTE THAT WE DONT GIVE THE ENTIRE PROTOCOL COMMAND TO THE EXTRACTION
    'PROCEEDURE. YOU NEEED TO OMIT THE FIRST FOUR CHARACTERS (PID), AS THAT
    'FIELD IS NOT REQUIRED BY THE EXTRACTION SUB.
    
    If modProtocol.ExtractProtData(Mid$(sThePacket, 5), strFields, 4) = False Then
            
        'The extract operation failed; the data is not the right format
        MsgBox "Data recieved is not in the required format.", 16
        
    Else
        
        Call ShowFields(strFields)
        
    End If
    
    
    If Not sPacketData = "" Then GoTo tcpbuffer
        
End Sub

Private Sub ShowIN(strString)
    
    txtOutput.Text = txtOutput.Text & "Data recieved:" & vbCrLf & strString & vbCrLf
    
End Sub

Private Sub ShowFields(strFields() As String)
    
    Dim strOut As String
    
    strOut = Join(strFields, vbCrLf)
    
    txtOutput.Text = txtOutput.Text & "Extracted Field Data:" & vbCrLf & strOut & vbCrLf
    
End Sub
 
Private Sub wsServer_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    
    txtOutput.Text = txtOutput.Text & "socket error:" & vbCrLf & Description & vbCrLf

End Sub
