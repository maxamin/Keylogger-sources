VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIPHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private m_ver           As Byte             'IP version number
Private m_len           As Byte             'header length in 32bit words (4 bits each)
Private m_tos           As Byte             'Type Of Service ID (1 octet)
Private m_totallength   As Integer          'Size of Datagram (header + data) in octets
Private m_id            As Long             'IP-ID (16 bits)
Private m_offset        As Integer          'fragmentation flags (3bit) and fragmet offset (13 bits)
Private m_ttl           As Byte             'datagram Time To Live (in network hops)
Private m_protocol      As Byte             'Transport protocol type (byte)
Private m_checksum      As Long             'Header Checksum (16 bits)
Private m_srcaddr       As Long             'Source IP Address (32 bits)
Private m_destaddr      As Long             'Destination IP Address (32 bits)
Private m_srcip         As String
Private m_destip        As String


Private Const DEFAULT_TTL As Long = 30



Public Property Get Version() As Byte
    Version = m_ver
End Property

Public Property Let Version(val As Byte)
    m_ver = val
End Property



Public Property Get HeaderLength() As Byte
    HeaderLength = m_len
End Property

Public Property Let HeaderLength(val As Byte)
    m_len = val
End Property



Public Property Get PacketLength() As Integer
    PacketLength = m_totallength
End Property

Public Property Let PacketLength(val As Integer)
    m_totallength = val
End Property


Public Property Get Checksum() As Long
    Checksum = m_checksum
End Property

Public Property Let Checksum(val As Long)
    m_checksum = val
End Property

Public Property Get ID() As Long
    ID = m_id
End Property

Public Property Let ID(val As Long)
    m_id = val
End Property



Public Property Get Offset() As Integer
    Offset = m_offset
End Property

Public Property Let Offset(val As Integer)
    m_offset = val
End Property



Public Property Get TimeToLive() As Byte
    TimeToLive = m_ttl
End Property

Public Property Let TimeToLive(val As Byte)
    m_ttl = val
End Property



Public Property Get Protocol() As IPProtocol
    Protocol = m_protocol
End Property

Public Property Let Protocol(val As IPProtocol)
    m_protocol = val
End Property



Public Property Get SourceAddress() As Long
    SourceAddress = m_srcaddr
End Property

Public Property Let SourceAddress(val As Long)
    m_srcaddr = val
    m_srcip = GetStrIPFromLong(val)
End Property



Public Property Get SourceIP() As String
    SourceIP = m_srcip
End Property

Public Property Let SourceIP(val As String)
    m_srcip = val
End Property


Public Property Get TypeOfService() As Long
    TypeOfService = m_tos
End Property

Public Property Let TypeOfService(val As Long)
    m_tos = val
End Property


Public Property Get DestAddress() As Long
    DestAddress = m_destaddr
End Property

Public Property Let DestAddress(val As Long)
    m_destaddr = val
    m_destip = GetStrIPFromLong(val)
End Property



Public Property Get DestIP() As String
    DestIP = m_destip
End Property

Public Property Let DestIP(val As String)
    m_destip = val
End Property
