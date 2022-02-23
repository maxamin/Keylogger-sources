VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsICMPProtocol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ICMPHeader
    icmp_type               As Byte
    icmp_code               As Byte
    icmp_checksum           As Integer
    icmp_ID                 As Long
End Type


Private m_icmp_type         As Byte
Private m_icmp_code         As Byte
Private m_icmp_checksum     As Long
Private m_icmp_ID           As Long

Public ProtocolInterface    As clsProtocolInterface

Public Event RecievedPacket(IPHeader As clsIPHeader, ICMPProtocol As clsICMPProtocol)


Public Sub PacketArrived(IPH As clsIPHeader, ReadBuffer() As Byte, BytesRecieved As Long)

  Dim ICMPH As ICMPHeader

    CopyMemory ICMPH, ByVal VarPtr(ReadBuffer(0)) + 20, LenB(ICMPH)

    Checksum = IntegerToUnsigned(ntohs(ICMPH.icmp_checksum))
    Code = ICMPH.icmp_code
    ID = ntohl(ICMPH.icmp_ID)
    TheType = ICMPH.icmp_type

    RaiseEvent RecievedPacket(IPH, Me)
    
End Sub



Public Property Get TheType() As ICMPType
    TheType = m_icmp_type
End Property

Public Property Let TheType(val As ICMPType)
    m_icmp_type = val
End Property


Public Property Get Code() As ICMPCodes
    Code = m_icmp_code
End Property

Public Property Let Code(val As ICMPCodes)
    m_icmp_code = val
End Property


Public Property Get Checksum() As Long
    Checksum = m_icmp_checksum
End Property

Public Property Let Checksum(val As Long)
    m_icmp_checksum = val
End Property


Public Property Get ID() As Long
    ID = m_icmp_ID
End Property

Public Property Let ID(val As Long)
    m_icmp_ID = val
End Property


Public Function GetICMPCodeStr() As String

    Select Case m_icmp_type
        Case ICMPType.[Destination Unreachable]
            Select Case m_icmp_code
                Case ICMPCodes.[Network Unreachable]: GetICMPCodeStr = "Network Unreachable"
                Case ICMPCodes.[Host Unreachable]: GetICMPCodeStr = "Host Unreachable"
                Case ICMPCodes.[Protocol Unreachable]: GetICMPCodeStr = "Protocol Unreachable"
                Case ICMPCodes.[Port Unreachable]: GetICMPCodeStr = "Port Unreachable"
                Case ICMPCodes.[Fragmentation Needed]: GetICMPCodeStr = "Fragmentation Needed"
            End Select
        Case ICMPType.Redirect
            Select Case m_icmp_code
                Case ICMPCodes.[Redirect Network]: GetICMPCodeStr = "Redirect Network"
                Case ICMPCodes.[Redirect Host]: GetICMPCodeStr = "Redirect Host"
                Case ICMPCodes.[Redirect TOS Network]: GetICMPCodeStr = "Redirect TOS Network"
                Case ICMPCodes.[Redirect TOS Host]: GetICMPCodeStr = "Redirect TOS Host"
            End Select
        Case ICMPType.[Time Exceeded]
            Select Case m_icmp_code
                Case ICMPCodes.[TTL Exceeded In Transit]: GetICMPCodeStr = "TTL Exceeded In Transit"
                Case ICMPCodes.[Reassembly Timeout]: GetICMPCodeStr = "Reassembly Timeout"
            End Select
        Case ICMPType.[Parameter Problem]
            Select Case m_icmp_code
                Case ICMPCodes.[Problem With Option]: GetICMPCodeStr = "Problem With Option"
            End Select
        Case Else
            GetICMPCodeStr = "N/A"
    End Select

End Function



Public Function GetICMPTypeStr() As String

    Select Case m_icmp_type
        Case ICMPType.[Echo Reply]: GetICMPTypeStr = "Echo Reply"
        Case 1, 2, 7: GetICMPTypeStr = "Reserved"
        Case ICMPType.[Destination Unreachable]: GetICMPTypeStr = "Destination Unreachable"
        Case ICMPType.[Source Quench]: GetICMPTypeStr = "Source Quench"
        Case ICMPType.Redirect: GetICMPTypeStr = "Redirect"
        Case ICMPType.[Alternate Host Address]: GetICMPTypeStr = "Alternate Host Address"
        Case ICMPType.[Echo Request]: GetICMPTypeStr = "Echo Request"
        Case ICMPType.[Router Advertisement]: GetICMPTypeStr = "Router Advertisement"
        Case ICMPType.[Router Solicitation]: GetICMPTypeStr = "Router Solicitation"
        Case ICMPType.[Time Exceeded]: GetICMPTypeStr = "Time Exceeded"
        Case ICMPType.[Parameter Problem]: GetICMPTypeStr = "Parameter Problem"
        Case ICMPType.[TimeStamp Reply]: GetICMPTypeStr = "Timestamp Request"
        Case ICMPType.[TimeStamp Request]: GetICMPTypeStr = "Timestamp Reply"
        Case ICMPType.[Information Request]: GetICMPTypeStr = "Information Request"
        Case ICMPType.[Information Reply]: GetICMPTypeStr = "Information Reply"
        Case ICMPType.[Address Mask Request]: GetICMPTypeStr = "Address Mask Request"
        Case ICMPType.[Address Mask Reply]: GetICMPTypeStr = "Address Mask Reply"
        Case 19: GetICMPTypeStr = "Reserved (security)"
        Case 20, 21, 22, 23, 24, 25, 26, 27, 28, 29: GetICMPTypeStr = "Reserved (robustness experiment)"
        Case ICMPType.[IP IX Trace Router]: GetICMPTypeStr = "TP IX Traceroute"
        Case ICMPType.[Conversion Error]: GetICMPTypeStr = "Conversion Error"
        Case ICMPType.[Mobile Host Redirect]: GetICMPTypeStr = "Mobile Host Redirect"
        Case ICMPType.[IPv6 Where Are You]: GetICMPTypeStr = "IPv6 Where Are You"
        Case ICMPType.[IPv6 Here I Am]: GetICMPTypeStr = "IPv6 I Am Here"
        Case ICMPType.[Mobile Registration Request]: GetICMPTypeStr = "Mobile Registration Request"
        Case ICMPType.[Mobile Registration Reply]: GetICMPTypeStr = "Mobile Registration Reply"
        Case ICMPType.[Domain Name Request]: GetICMPTypeStr = "Domain Name Request"
        Case ICMPType.[Domain Name Reply]: GetICMPTypeStr = "Domain Name Reply"
        Case ICMPType.[SKIP Algorithm Discovery Protocol]: GetICMPTypeStr = "SKIP Algorithm Discovery_protocol"
        Case ICMPType.[IPsec Security Failures]: GetICMPTypeStr = "IPsec Security Failures"
        Case Is > 40: GetICMPTypeStr = "Reserved"
    End Select

End Function

