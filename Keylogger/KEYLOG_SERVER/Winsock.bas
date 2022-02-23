Attribute VB_Name = "modWinsock"
Option Explicit

'Select uses arrays of SOCKETs.  These macros manipulate such
'arrays.  FD_SETSIZE may be defined by the user before including
'this file, but the default here should be >= 64.

Public Const FD_SETSIZE         As Long = 64
Public Const FD_MAX_EVENTS      As Long = 10

Public Type fd_set
  fd_count                      As Long                     'How many are SET?
  fd_array(1 To FD_SETSIZE)     As Long                     'An array of SOCKETs
End Type


'Structure used in select() call, taken from the BSD file sys/time.h.
Public Type timeval
    tv_sec                      As Long                     'Seconds
    tv_usec                     As Long                     'And microseconds
End Type


' Commands for ioctlsocket(),  taken from the BSD file fcntl.h.
'
' Ioctl's have the command encoded in the lower word,
' and the size of any in or out parameters in the upper
' word.  The high 2 bits of the upper word are used
' to encode the in/out status of the parameter; for now
' we restrict parameters to at most 128 bytes.

Public Const IOCPARM_MASK       As Long = &H7F              'Parameters must be < 128 bytes
Public Const IOC_VOID           As Long = &H20000000        'No parameters
Public Const IOC_OUT            As Long = &H40000000        'Copy out parameters
Public Const IOC_IN             As Long = &H80000000        'Copy in parameters
Public Const IOC_INOUT          As Long = (IOC_IN Or IOC_OUT)

'Structures returned by network data base library, taken from the
'BSD file netdb.h.  All addresses are supplied in host order, and
'returned in network order (suitable for use in system calls).

Public Type HostEnt
    h_name                      As Long
    h_aliases                   As Long
    h_addrtype                  As Integer
    h_length                    As Integer
    h_addr_list                 As Long
End Type


Public Type servent
    s_name                      As Long
    s_aliases                   As Long
    s_port                      As Integer
    s_proto                     As Long
End Type


Public Type protoent
    p_name                      As String                   'Official name of the protocol
    p_aliases                   As Long                     'Null-terminated array of alternate names
    p_proto                     As Long                     'Protocol number, in host byte order
End Type


'Constants and structures defined by the internet system,
'Per RFC 790, September 1981, taken from the BSD file netinet/in.h.

'Protocols
Public Enum IPProtocol
             IPPROTO_IP = 0                                 'Dummy for IP
             IPPROTO_ICMP = 1                               'Control message protocol
             IPPROTO_IGMP = 2                               'Group management protocol
             IPPROTO_GGP = 3                                'Gateway^2 (deprecated)
             IPPROTO_TCP = 6                                'Tcp
             IPPROTO_PUP = 12                               'Pup
             IPPROTO_UDP = 17                               'User datagram protocol
             IPPROTO_IDP = 22                               'Xns idp
             IPPROTO_ND = 77                                'UNOFFICIAL net disk proto
End Enum


Public Const IPPROTO_RAW = 255                              'Raw IP packet
Public Const IPPROTO_MAX = 256

'Port/socket numbers: network standard functions

Public Enum ServicePort
             IPPORT_ECHO = 7
             IPPORT_DISCARD = 9
             IPPORT_SYSTAT = 11
             IPPORT_DAYTIME = 13
             IPPORT_NETSTAT = 15
             IPPORT_FTP = 21
             IPPORT_TELNET = 23
             IPPORT_SMTP = 25
             IPPORT_TIMESERVER = 37
             IPPORT_NAMESERVER = 42
             IPPORT_WHOIS = 43
             IPPORT_MTP = 57
End Enum

'Port/socket numbers: host specific functions
Public Const IPPORT_TFTP        As Long = 69
Public Const IPPORT_RJE         As Long = 77
Public Const IPPORT_FINGER      As Long = 79
Public Const IPPORT_TTYLINK     As Long = 87
Public Const IPPORT_SUPDUP      As Long = 95

'UNIX TCP sockets
Public Const IPPORT_EXECSERVER  As Long = 512
Public Const IPPORT_LOGINSERVER As Long = 513
Public Const IPPORT_CMDSERVER   As Long = 514
Public Const IPPORT_EFSSERVER   As Long = 520

'UNIX UDP sockets
Public Const IPPORT_BIFFUDP     As Long = 512
Public Const IPPORT_WHOSERVER   As Long = 513
Public Const IPPORT_ROUTESERVER As Long = 520


'Ports < IPPORT_RESERVED are reserved for
'privileged processes (e.g. root).
Public Const IPPORT_RESERVED    As Long = 1024

'Link numbers
Public Const IMPLINK_IP         As Long = 155
Public Const IMPLINK_LOWEXPER   As Long = 156
Public Const IMPLINK_HIGHEXPER  As Long = 158



'Definitions of bits in internet address integers.
'On subnets, the decomposition of addresses to host and net parts
'is done according to subnet mask, not the masks here.
Public Const IN_CLASSA_NET      As Long = &HFF000000
Public Const IN_CLASSA_NSHIFT   As Long = 24
Public Const IN_CLASSA_HOST     As Long = &HFFFFFF
Public Const IN_CLASSA_MAX      As Long = 128

Public Const IN_CLASSB_NET      As Long = &HFFFF0000
Public Const IN_CLASSB_NSHIFT   As Long = 16
Public Const IN_CLASSB_HOST     As Long = &HFFFF
Public Const IN_CLASSB_MAX      As Long = 65536

Public Const IN_CLASSC_NET      As Long = &HFFFFFF00
Public Const IN_CLASSC_NSHIFT   As Long = 8
Public Const IN_CLASSC_HOST     As Long = &HFF

Public Const INADDR_LOOPBACK    As Long = &H7F000001
Public Const INADDR_BROADCAST   As Long = &HFFFFFFFF
Public Const INADDR_NONE        As Long = &HFFFFFFFF

'Socket address, internet style.
Public Type sockaddr_in
    sin_family                  As Integer
    sin_port                    As Integer
    sin_addr                    As Long
    sin_zero                    As String * 8
End Type

Public Const WSADESCRIPTION_LEN As Long = 256
Public Const WSASYS_STATUS_LEN  As Long = 128

Public Type WSAData
    wVersion                    As Integer
    wHighVersion                As Integer
    szDescription               As String * WSADESCRIPTION_LEN
    szSystemStatus              As String * WSASYS_STATUS_LEN
    iMaxSockets                 As Integer
    iMaxUdpDg                   As Integer
    lpVendorInfo                As Long
End Type

'Options for use with (gs)etsockopt at the IP level.
Public Const IP_OPTIONS         As Long = 1                 'Set/get IP per-packet options
Public Const IP_MULTICAST_IF    As Long = 2                 'Set/get IP multicast interface
Public Const IP_MULTICAST_TTL   As Long = 3                 'Set/get IP multicast timetolive
Public Const IP_MULTICAST_LOOP  As Long = 4                 'Set/get IP multicast loopback
Public Const IP_ADD_MEMBERSHIP  As Long = 5                 'Add an IP group membership
Public Const IP_DROP_MEMBERSHIP As Long = 6                 'Drop an IP group membership
Public Const ip_ttl             As Long = 7                 'Set/get IP Time To Live
Public Const ip_tos             As Long = 8                 'Set/get IP Type Of Service
Public Const IP_DONTFRAGMENT    As Long = 9                 'Set/get IP Don't Fragment flag
    
    
Public Const IP_DEFAULT_MULTICAST_TTL   As Long = 1         'normally limit m'casts to 1 hop
Public Const IP_DEFAULT_MULTICAST_LOOP  As Long = 1         'normally hear sends if a member
Public Const IP_MAX_MEMBERSHIPS         As Long = 20        'per socket '* must fit in one mbuf


'This is used instead of -1, since the
'SOCKET type is unsigned.
Public Const INVALID_SOCKET     As Long = 0
Public Const SOCKET_ERROR       As Long = -1

'Types
Public Const SOCK_STREAM        As Long = 1                 'Stream socket
Public Const SOCK_DGRAM         As Long = 2                 'Datagram socket
Public Const SOCK_RAW           As Long = 3                 'Raw-protocol interface
Public Const SOCK_RDM           As Long = 4                 'Reliably-delivered message
Public Const SOCK_SEQPACKET     As Long = 5                 'Sequenced packet stream

'Option flags per-socket.
Public Const SO_DEBUG           As Long = &H1               'Turn on debugging info recording
Public Const SO_ACCEPTCONN      As Long = &H2               'Socket has had listen()
Public Const SO_REUSEADDR       As Long = &H4               'Allow local address reuse
Public Const SO_KEEPALIVE       As Long = &H8               'Keep connections alive
Public Const SO_DONTROUTE       As Long = &H10              'Just use interface addresses
Public Const SO_BROADCAST       As Long = &H20              'Permit sending of broadcast msgs
Public Const SO_USELOOPBACK     As Long = &H40              'Bypass hardware when possible
Public Const SO_LINGER          As Long = &H80              'Linger on close if data present
Public Const SO_OOBINLINE       As Long = &H100             'Leave received OOB data in line

Public Const SO_DONTLINGER      As Long = Not SO_LINGER

'Additional options.
Public Const SO_SNDBUF          As Long = &H1001            'Send buffer size
Public Const SO_RCVBUF          As Long = &H1002            'Receive buffer size
Public Const SO_SNDLOWAT        As Long = &H1003            'Send low-water mark
Public Const SO_RCVLOWAT        As Long = &H1004            'Receive low-water mark
Public Const SO_SNDTIMEO        As Long = &H1005            'Send timeout
Public Const SO_RCVTIMEO        As Long = &H1006            'Receive timeout
Public Const SO_ERROR           As Long = &H1007            'Get error status and clear
Public Const SO_TYPE            As Long = &H1008            'Get socket type


'WinSock 2 extension -- new options
Public Const SO_GROUP_ID        As Long = &H2001            'ID of a socket group
Public Const SO_GROUP_PRIORITY  As Long = &H2002            'The relative priority within a group
Public Const SO_MAX_MSG_SIZE    As Long = &H2003            'maximum message size
Public Const SO_PROTOCOL_INFOA  As Long = &H2004            'WSAPROTOCOL_INFOA structure
Public Const SO_PROTOCOL_INFOW  As Long = &H2005            'WSAPROTOCOL_INFOW structure
Public Const SO_PROTOCOL_INFO   As Long = SO_PROTOCOL_INFOW
Public Const PVD_CONFIG         As Long = &H3001            'Configuration info for service provider


'Options for connect and disconnect data and options. Used only by
'non-TCP/IP transports such as DECNet, OSI TP4, etc.
Public Const SO_CONNDATA        As Long = &H7000
Public Const SO_CONNOPT         As Long = &H7001
Public Const SO_DISCDATA        As Long = &H7002
Public Const SO_DISCOPT         As Long = &H7003
Public Const SO_CONNDATALEN     As Long = &H7004
Public Const SO_CONNOPTLEN      As Long = &H7005
Public Const SO_DISCDATALEN     As Long = &H7006
Public Const SO_DISCOPTLEN      As Long = &H7007

' Option for opening sockets for synchronous access.
Public Const SO_OPENTYPE        As Long = &H7008

Public Const SO_SYNCHRONOUS_ALERT       As Long = &H10
Public Const SO_SYNCHRONOUS_NONALERT    As Long = &H20

' Other NT-specific options.
Public Const SO_MAXDG                   As Long = &H7009
Public Const SO_MAXPATHDG               As Long = &H700A
Public Const SO_UPDATE_ACCEPT_CONTEXT   As Long = &H700B
Public Const SO_CONNECT_TIME            As Long = &H700C

' TCP options.
Public Const TCP_NODELAY        As Long = &H1
Public Const TCP_BSDURGENT      As Long = &H7000

' Address families.
Public Const AF_UNSPEC          As Long = 0                 'Unspecified
Public Const AF_UNIX            As Long = 1                 'Local to host (pipes, portals)
Public Const AF_INET            As Long = 2                 'Internetwork: UDP, TCP, etc.
Public Const AF_IMPLINK         As Long = 3                 'Arpanet imp addresses
Public Const AF_PUP             As Long = 4                 'Pup protocols: e.g. BSP
Public Const AF_CHAOS           As Long = 5                 'Mit CHAOS protocols
Public Const AF_IPX             As Long = 6                 'IPX and SPX
Public Const AF_NS              As Long = 6                 'XEROX NS protocols
Public Const AF_ISO             As Long = 7                 'ISO protocols
Public Const AF_OSI             As Long = AF_ISO            'OSI is ISO
Public Const AF_ECMA            As Long = 8                 'European computer manufacturers
Public Const AF_DATAKIT         As Long = 9                 'Datakit protocols
Public Const AF_CCITT           As Long = 10                'CCITT protocols, X.25 etc
Public Const AF_SNA             As Long = 11                'IBM SNA
Public Const AF_DECnet          As Long = 12                'DECnet
Public Const AF_DLI             As Long = 13                'Direct data link interface
Public Const AF_LAT             As Long = 14                'LAT
Public Const AF_HYLINK          As Long = 15                'NSC Hyperchannel
Public Const AF_APPLETALK       As Long = 16                'AppleTalk
Public Const AF_NETBIOS         As Long = 17                'NetBios-style addresses
Public Const AF_VOICEVIEW       As Long = 18                'VoiceView
Public Const AF_FIREFOX         As Long = 19                'FireFox
Public Const AF_UNKNOWN1        As Long = 20                'Somebody is using this!
Public Const AF_BAN             As Long = 21                'Banyan

Public Const AF_MAX             As Long = 22

'Protocol families, same as address families for now.
Public Const PF_UNSPEC          As Long = AF_UNSPEC
Public Const PF_UNIX            As Long = AF_UNIX
Public Const PF_INET            As Long = AF_INET
Public Const PF_IMPLINK         As Long = AF_IMPLINK
Public Const PF_PUP             As Long = AF_PUP
Public Const PF_CHAOS           As Long = AF_CHAOS
Public Const PF_NS              As Long = AF_NS
Public Const PF_IPX             As Long = AF_IPX
Public Const PF_ISO             As Long = AF_ISO
Public Const PF_OSI             As Long = AF_OSI
Public Const PF_ECMA            As Long = AF_ECMA
Public Const PF_DATAKIT         As Long = AF_DATAKIT
Public Const PF_CCITT           As Long = AF_CCITT
Public Const PF_SNA             As Long = AF_SNA
Public Const PF_DECnet          As Long = AF_DECnet
Public Const PF_DLI             As Long = AF_DLI
Public Const PF_LAT             As Long = AF_LAT
Public Const PF_HYLINK          As Long = AF_HYLINK
Public Const PF_APPLETALK       As Long = AF_APPLETALK
Public Const PF_VOICEVIEW       As Long = AF_VOICEVIEW
Public Const PF_FIREFOX         As Long = AF_FIREFOX
Public Const PF_UNKNOWN1        As Long = AF_UNKNOWN1
Public Const PF_BAN             As Long = AF_BAN

Public Const PF_MAX             As Long = AF_MAX


'Structure used for manipulating linger option.
Public Type LingerType
    l_onoff                     As Integer
    l_linger                    As Integer
End Type


'Input Output option
Public Const SIO_RCVALL         As Long = &H98000001

'Level number for (get/set)sockopt() to apply to socket itself.
Public Const SOL_SOCKET         As Long = &HFFFF&           'Options for socket level

'Maximum queue length specifiable by listen.
Public Const SOMAXCONN          As Long = 5

Public Const MSG_OOB            As Long = &H1               'Process out-of-band data
Public Const MSG_PEEK           As Long = &H2               'Peek at incoming message
Public Const MSG_DONTROUTE      As Long = &H4               'Send without using routing tables

Public Const MSG_MAXIOVLEN      As Long = 16

Public Const MSG_PARTIAL        As Long = &H8000            'Partial send or recv for message xport


'WinSock 2 extension -- new flags for WSASend(), WSASendTo(), WSARecv() and
'                                     WSARecvFrom()
Public Const MSG_INTERRUPT      As Long = &H10              'Send/recv in the interrupt context


'Define constant based on rfc883, used by gethostbyxxxx() calls.
Public Const MAXGETHOSTSTRUCT   As Long = 1024

'Define flags to be used with the WSAAsyncSelect() call.
Public Const FD_READ            As Long = &H1
Public Const FD_WRITE           As Long = &H2
Public Const FD_OOB             As Long = &H4
Public Const FD_ACCEPT          As Long = &H8
Public Const FD_CONNECT         As Long = &H10
Public Const FD_CLOSE           As Long = &H20

'All Windows Sockets error constants are biased by WSABASEERR from
'the "normal"
Public Const WSABASEERR         As Long = 10000

'Windows Sockets definitions of regular Microsoft C error constants
Public Const WSAEINTR           As Long = (WSABASEERR + 4)
Public Const WSAEBADF           As Long = (WSABASEERR + 9)
Public Const WSAEACCES          As Long = (WSABASEERR + 13)
Public Const WSAEFAULT          As Long = (WSABASEERR + 14)
Public Const WSAEINVAL          As Long = (WSABASEERR + 22)
Public Const WSAEMFILE          As Long = (WSABASEERR + 24)

'Windows Sockets definitions of regular Berkeley error constants
Public Const WSAEWOULDBLOCK     As Long = (WSABASEERR + 35)
Public Const WSAEINPROGRESS     As Long = (WSABASEERR + 36)
Public Const WSAEALREADY        As Long = (WSABASEERR + 37)
Public Const WSAENOTSOCK        As Long = (WSABASEERR + 38)
Public Const WSAEDESTADDRREQ    As Long = (WSABASEERR + 39)
Public Const WSAEMSGSIZE        As Long = (WSABASEERR + 40)
Public Const WSAEPROTOTYPE      As Long = (WSABASEERR + 41)
Public Const WSAENOPROTOOPT     As Long = (WSABASEERR + 42)
Public Const WSAEPROTONOSUPPORT As Long = (WSABASEERR + 43)
Public Const WSAESOCKTNOSUPPORT As Long = (WSABASEERR + 44)
Public Const WSAEOPNOTSUPP      As Long = (WSABASEERR + 45)
Public Const WSAEPFNOSUPPORT    As Long = (WSABASEERR + 46)
Public Const WSAEAFNOSUPPORT    As Long = (WSABASEERR + 47)
Public Const WSAEADDRINUSE      As Long = (WSABASEERR + 48)
Public Const WSAEADDRNOTAVAIL   As Long = (WSABASEERR + 49)
Public Const WSAENETDOWN        As Long = (WSABASEERR + 50)
Public Const WSAENETUNREACH     As Long = (WSABASEERR + 51)
Public Const WSAENETRESET       As Long = (WSABASEERR + 52)
Public Const WSAECONNABORTED    As Long = (WSABASEERR + 53)
Public Const WSAECONNRESET      As Long = (WSABASEERR + 54)
Public Const WSAENOBUFS         As Long = (WSABASEERR + 55)
Public Const WSAEISCONN         As Long = (WSABASEERR + 56)
Public Const WSAENOTCONN        As Long = (WSABASEERR + 57)
Public Const WSAESHUTDOWN       As Long = (WSABASEERR + 58)
Public Const WSAETOOMANYREFS    As Long = (WSABASEERR + 59)
Public Const WSAETIMEDOUT       As Long = (WSABASEERR + 60)
Public Const WSAECONNREFUSED    As Long = (WSABASEERR + 61)
Public Const WSAELOOP           As Long = (WSABASEERR + 62)
Public Const WSAENAMETOOLONG    As Long = (WSABASEERR + 63)
Public Const WSAEHOSTDOWN       As Long = (WSABASEERR + 64)
Public Const WSAEHOSTUNREACH    As Long = (WSABASEERR + 65)
Public Const WSAENOTEMPTY       As Long = (WSABASEERR + 66)
Public Const WSAEPROCLIM        As Long = (WSABASEERR + 67)
Public Const WSAEUSERS          As Long = (WSABASEERR + 68)
Public Const WSAEDQUOT          As Long = (WSABASEERR + 69)
Public Const WSAESTALE          As Long = (WSABASEERR + 70)
Public Const WSAEREMOTE         As Long = (WSABASEERR + 71)

Public Const WSAEDISCON         As Long = (WSABASEERR + 101)

'WinSock 2 extension -- new error codes and type definition
Public Const WSA_IO_PENDING         As Long = (WSAEWOULDBLOCK)
Public Const WSA_IO_INCOMPLETE      As Long = (WSAEWOULDBLOCK)
Public Const WSA_INVALID_HANDLE     As Long = (WSAENOTSOCK)
Public Const WSA_INVALID_PARAMETER  As Long = (WSAEINVAL)
Public Const WSA_NOT_ENOUGH_MEMORY  As Long = (WSAENOBUFS)
Public Const WSA_OPERATION_ABORTED  As Long = (WSAEINTR)


'Extended Windows Sockets error constant definitions
Public Const WSASYSNOTREADY     As Long = (WSABASEERR + 91)
Public Const WSAVERNOTSUPPORTED As Long = (WSABASEERR + 92)
Public Const WSANOTINITIALISED  As Long = (WSABASEERR + 93)


'Authoritative Answer: Host not found
Public Const WSAHOST_NOT_FOUND  As Long = (WSABASEERR + 1001)
Public Const HOST_NOT_FOUND     As Long = WSAHOST_NOT_FOUND

'Non-Authoritative: Host not found, or SERVERFAIL
Public Const WSATRY_AGAIN       As Long = (WSABASEERR + 1002)
Public Const TRY_AGAIN          As Long = WSATRY_AGAIN

'Non recoverable errors, FORMERR, REFUSED, NOTIMP
Public Const WSANO_RECOVERY     As Long = (WSABASEERR + 1003)
Public Const NO_RECOVERY        As Long = WSANO_RECOVERY

'Valid name, no data record of requested type
Public Const WSANO_DATA         As Long = (WSABASEERR + 1004)
Public Const NO_DATA            As Long = WSANO_DATA

'no address, look for MX record
Public Const WSANO_ADDRESS      As Long = WSANO_DATA
Public Const NO_ADDRESS         As Long = WSANO_ADDRESS

'Windows Sockets errors redefined as regular Berkeley error constants.
'These are commented out in Windows NT to avoid conflicts with errno.h.
'Use the WSA constants instead.
Public Const EWOULDBLOCK        As Long = WSAEWOULDBLOCK
Public Const EINPROGRESS        As Long = WSAEINPROGRESS
Public Const EALREADY           As Long = WSAEALREADY
Public Const ENOTSOCK           As Long = WSAENOTSOCK
Public Const EDESTADDRREQ       As Long = WSAEDESTADDRREQ
Public Const EMSGSIZE           As Long = WSAEMSGSIZE
Public Const EPROTOTYPE         As Long = WSAEPROTOTYPE
Public Const ENOPROTOOPT        As Long = WSAENOPROTOOPT
Public Const EPROTONOSUPPORT    As Long = WSAEPROTONOSUPPORT
Public Const ESOCKTNOSUPPORT    As Long = WSAESOCKTNOSUPPORT
Public Const EOPNOTSUPP         As Long = WSAEOPNOTSUPP
Public Const EPFNOSUPPORT       As Long = WSAEPFNOSUPPORT
Public Const EAFNOSUPPORT       As Long = WSAEAFNOSUPPORT
Public Const EADDRINUSE         As Long = WSAEADDRINUSE
Public Const EADDRNOTAVAIL      As Long = WSAEADDRNOTAVAIL
Public Const ENETDOWN           As Long = WSAENETDOWN
Public Const ENETUNREACH        As Long = WSAENETUNREACH
Public Const ENETRESET          As Long = WSAENETRESET
Public Const ECONNABORTED       As Long = WSAECONNABORTED
Public Const ECONNRESET         As Long = WSAECONNRESET
Public Const ENOBUFS            As Long = WSAENOBUFS
Public Const EISCONN            As Long = WSAEISCONN
Public Const ENOTCONN           As Long = WSAENOTCONN
Public Const ESHUTDOWN          As Long = WSAESHUTDOWN
Public Const ETOOMANYREFS       As Long = WSAETOOMANYREFS
Public Const ETIMEDOUT          As Long = WSAETIMEDOUT
Public Const ECONNREFUSED       As Long = WSAECONNREFUSED
Public Const ELOOP              As Long = WSAELOOP
Public Const ENAMETOOLONG       As Long = WSAENAMETOOLONG
Public Const EHOSTDOWN          As Long = WSAEHOSTDOWN
Public Const EHOSTUNREACH       As Long = WSAEHOSTUNREACH
Public Const ENOTEMPTY          As Long = WSAENOTEMPTY
Public Const EPROCLIM           As Long = WSAEPROCLIM
Public Const EUSERS             As Long = WSAEUSERS
Public Const EDQUOT             As Long = WSAEDQUOT
Public Const ESTALE             As Long = WSAESTALE
Public Const EREMOTE            As Long = WSAEREMOTE


Public Const TF_DISCONNECT      As Long = &H1
Public Const TF_REUSE_SOCKET    As Long = &H2
Public Const TF_WRITE_BEHIND    As Long = &H4


'WinSock 2 extension -- manifest constants for return values of the condition function
Public Const CF_ACCEPT          As Long = &H0
Public Const CF_REJECT          As Long = &H1
Public Const CF_DEFER           As Long = &H2

'WinSock 2 extension -- manifest constants for shutdown()
Public Const SD_RECEIVE         As Long = &H0
Public Const SD_SEND            As Long = &H1
Public Const SD_BOTH            As Long = &H2


Public Type WSANETWORKEVENTS
    lNetworkEvents              As Long
    iErrorCode(FD_MAX_EVENTS)   As Long
End Type


Public Const MAX_PROTOCOL_CHAIN = 7
    
Public Const BASE_PROTOCOL = 1
Public Const LAYERED_PROTOCOL = 0
    
Public Type WSAPROTOCOLCHAIN
    ChainLen                    As Long                     'The length of the chain,
                                                            'Length = 0 means layered protocol,
                                                            'Length = 1 means base protocol,
                                                            'Length > 1 means protocol chain
    ChainEntries(MAX_PROTOCOL_CHAIN) As Long                'A list of dwCatalogEntryIds
End Type


Public Const WSAPROTOCOL_LEN    As Long = 255

Public Type WSAPROTOCOL_INFOA
    dwServiceFlags1             As Long
    dwServiceFlags2             As Long
    dwServiceFlags3             As Long
    dwServiceFlags4             As Long
    dwProviderFlags             As Long
    ProviderId                  As String
    dwCatalogEntryId            As Long
    WSAPROTOCOLCHAIN            As WSAPROTOCOLCHAIN
    iVersion                    As Long
    iAddressFamily              As Long
    iMaxSockAddr                As Long
    iMinSockAddr                As Long
    iSocketType                 As Long
    iProtocol                   As Long
    iProtocolMaxOffset          As Long
    iNetworkByteOrder           As Long
    iSecurityScheme             As Long
    dwMessageSize               As Long
    dwProviderReserved          As Long
    szProtocol(WSAPROTOCOL_LEN + 1) As String * 1
End Type


Public Type WSAPROTOCOL_INFOW
    dwServiceFlags1             As Long
    dwServiceFlags2             As Long
    dwServiceFlags3             As Long
    dwServiceFlags4             As Long
    dwProviderFlags             As Long
    ProviderId                  As String
    dwCatalogEntryId            As Long
    ProtocolChain               As WSAPROTOCOLCHAIN
    iVersion                    As Long
    iAddressFamily              As Long
    iMaxSockAddr                As Long
    iMinSockAddr                As Long
    iSocketType                 As Long
    iProtocol                   As Long
    iProtocolMaxOffset          As Long
    iNetworkByteOrder           As Long
    iSecurityScheme             As Long
    dwMessageSize               As Long
    dwProviderReserved          As Long
    szProtocol(WSAPROTOCOL_LEN + 1) As String * 1
End Type



'Flag bit definitions for dwProviderFlags
Public Const PFL_MULTIPLE_PROTO_ENTRIES     As Long = &H1
Public Const PFL_RECOMMENDED_PROTO_ENTRY    As Long = &H2
Public Const PFL_HIDDEN                     As Long = &H4
Public Const PFL_MATCHES_PROTOCOL_ZERO      As Long = &H8

'Flag bit definitions for dwServiceFlags1
Public Const XP1_CONNECTIONLESS             As Long = &H1
Public Const XP1_GUARANTEED_DELIVERY        As Long = &H2
Public Const XP1_GUARANTEED_ORDER           As Long = &H4
Public Const XP1_MESSAGE_ORIENTED           As Long = &H8
Public Const XP1_PSEUDO_STREAM              As Long = &H10
Public Const XP1_GRACEFUL_CLOSE             As Long = &H20
Public Const XP1_EXPEDITED_DATA             As Long = &H40
Public Const XP1_CONNECT_DATA               As Long = &H80
Public Const XP1_DISCONNECT_DATA            As Long = &H100
Public Const XP1_SUPPORT_BROADCAST          As Long = &H200
Public Const XP1_SUPPORT_MULTIPOINT         As Long = &H400
Public Const XP1_MULTIPOINT_CONTROL_PLANE   As Long = &H800
Public Const XP1_MULTIPOINT_DATA_PLANE      As Long = &H1000
Public Const XP1_QOS_SUPPORTED              As Long = &H2000
Public Const XP1_INTERRUPT                  As Long = &H4000
Public Const XP1_UNI_SEND                   As Long = &H8000
Public Const XP1_UNI_RECV                   As Long = &H10000
Public Const XP1_IFS_HANDLES                As Long = &H20000
Public Const XP1_PARTIAL_MESSAGE            As Long = &H40000

Public Const BIGENDIAN                      As Long = &H0
Public Const LITTLEENDIAN                   As Long = &H1

Public Const SECURITY_PROTOCOL_NONE         As Long = &H0

'WinSock 2 extension -- manifest constants for WSAJoinLeaf()
Public Const JL_SENDER_ONLY             As Long = &H1
Public Const JL_RECEIVER_ONLY           As Long = &H2
Public Const JL_BOTH                    As Long = &H4

'WinSock 2 extension -- manifest constants for WSASocket()
Public Const WSA_FLAG_OVERLAPPED        As Long = &H1
Public Const WSA_FLAG_MULTIPOINT_C_ROOT As Long = &H2
Public Const WSA_FLAG_MULTIPOINT_C_LEAF As Long = &H4
Public Const WSA_FLAG_MULTIPOINT_D_ROOT As Long = &H8
Public Const WSA_FLAG_MULTIPOINT_D_LEAF As Long = &H10

'WinSock 2 extension -- manifest constants for WSAIoctl()
Public Const IOC_UNIX                   As Long = &H0
Public Const IOC_WS2                    As Long = &H8000000
Public Const IOC_PROTOCOL               As Long = &H10000000
Public Const IOC_VENDOR                 As Long = &H18000000


'WinSock 2 extension -- manifest constants for SIO_TRANSLATE_HANDLE ioctl
Public Const TH_NETDEV                  As Long = &H1
Public Const TH_TAPI                    As Long = &H2


'Resolution flags for WSAGetAddressByName().
'Note these are also used by the 1.1 API GetAddressByName, so
'leave them around.
Public Const RES_UNUSED_1               As Long = &H1
Public Const RES_FLUSH_CACHE            As Long = &H2
Public Const RES_SERVICE                As Long = &H4


Public Type OVERLAPPED
    Internal                            As Long
    InternalHigh                        As Long
    Offset                              As Long
    OffsetHigh                          As Long
    hEvent                              As Long
End Type


Public Enum ICMPType
        [Echo Reply] = 0
        [Destination Unreachable] = 3
        [Source Quench] = 4
        [Redirect] = 5
        [Alternate Host Address] = 6
        [Echo Request] = 8
        [Router Advertisement] = 9
        [Router Solicitation] = 10
        [Time Exceeded] = 11
        [Parameter Problem] = 12
        [TimeStamp Request] = 13
        [TimeStamp Reply] = 14
        [Information Request] = 15
        [Information Reply] = 16
        [Address Mask Request] = 17
        [Address Mask Reply] = 18
        [IP IX Trace Router] = 30
        [Conversion Error] = 31
        [Mobile Host Redirect] = 32
        [IPv6 Where Are You] = 33
        [IPv6 Here I Am] = 34
        [Mobile Registration Request] = 35
        [Mobile Registration Reply] = 36
        [Domain Name Request] = 37
        [Domain Name Reply] = 38
        [SKIP Algorithm Discovery Protocol] = 39
        [IPsec Security Failures] = 40
End Enum


Public Enum ICMPCodes
        [Network Unreachable] = 0
        [Host Unreachable] = 1
        [Protocol Unreachable] = 2
        [Port Unreachable] = 3
        [Fragmentation Needed] = 4

        [Redirect Network] = 0
        [Redirect Host] = 1
        [Redirect TOS Network] = 2
        [Redirect TOS Host] = 3

        [TTL Exceeded In Transit] = 0
        [Reassembly Timeout] = 1

        [Problem With Option] = 0
End Enum


Public Const SIO_GET_INTERFACE_LIST     As Long = &H4004747F

Public Type sockaddr_gen
    AddressIn                           As sockaddr_in
    filler(0 To 7)                      As Byte
End Type
  
  
Public Type INTERFACE_INFO
    iiFlags                             As Long             'Interface flags
    iiAddress                           As sockaddr_gen     'Interface address
    iiBroadcastAddress                  As sockaddr_gen     'Broadcast address
    iiNetmask                           As sockaddr_gen     'Network mask
End Type


Public Type INTERFACEINFO
    iInfo(0 To 7) As INTERFACE_INFO
End Type


Public Const WSA_INFINITE               As Long = -1


Public Type IPHeader
    ip_verlen                           As Byte             'IP version number and header length in 32bit words (4 bits each)
    ip_tos                              As Byte             'Type Of Service ID (1 octet)
    ip_totallength                      As Integer          'Size of Datagram (header + data) in octets
    ip_id                               As Integer          'IP-ID (16 bits)
    ip_offset                           As Integer          'fragmentation flags (3bit) and fragmet offset (13 bits)
    ip_ttl                              As Byte             'datagram Time To Live (in network hops)
    ip_protocol                         As Byte             'Transport protocol type (byte)
    ip_checksum                         As Integer          'Header Checksum (16 bits)
    ip_srcaddr                          As Long             'Source IP Address (32 bits)
    ip_destaddr                         As Long             'Destination IP Address (32 bits)
End Type


'Socket startup and cleanuo
Public Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Public Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Public Declare Function shutdown Lib "wsock32" (ByVal hSocket As Long, ByVal nHow As Long) As Long


'Creating and closing sockets
Public Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal s_type As Long, ByVal Protocol As Long) As Long
Public Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long

'Connecting and listening for connections
Public Declare Function Connect Lib "ws2_32.dll" Alias "connect" (ByVal s As Long, ByRef name As sockaddr_in, ByVal namelen As Long) As Long
Public Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Public Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As sockaddr_in, ByRef addrlen As Long) As Long
Public Declare Function AcceptEx Lib "Mswsock.dll" (ByVal sListenSocket As Long, ByVal sAcceptSocket As Long, lpOutputBuffer As Byte, ByVal dwReceiveDataLength As Long, ByVal dwLocalAddressLength As Long, ByVal dwRemoteAddressLength As Long, lpdwBytesReceived As Long, lpOverlapped As OVERLAPPED) As Long

'Database functions
Public Declare Function gethostbyaddr Lib "ws2_32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
Public Declare Function gethostbyname Lib "ws2_32.dll" (ByVal host_name As String) As Long
Public Declare Function GetHostName Lib "ws2_32.dll" Alias "gethostname" (ByVal host_name As String, ByVal namelen As Long) As Long
Public Declare Function getservbyname Lib "ws2_32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
Public Declare Function getprotobynumber Lib "ws2_32.dll" (ByVal proto As Long) As Long
Public Declare Function getprotobyname Lib "ws2_32.dll" (ByVal proto_name As String) As Long
Public Declare Function getservbyport Lib "ws2_32.dll" (ByVal Port As Integer, ByVal proto As Long) As Long

Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal number As Long, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long


'Conversion functions
Public Declare Function inet_addr Lib "ws2_32.dll" (ByVal cp As String) As Long
Public Declare Function inet_ntoa Lib "ws2_32.dll" (ByVal inn As Long) As Long
Public Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Public Declare Function htonl Lib "ws2_32.dll" (ByVal hostlong As Long) As Long
Public Declare Function ntohl Lib "ws2_32.dll" (ByVal netlong As Long) As Long
Public Declare Function ntohs Lib "ws2_32.dll" (ByVal netshort As Integer) As Integer
    
'Socket information
Public Declare Function getsockname Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function getpeername Lib "ws2_32.dll" (ByVal s As Long, ByRef name As sockaddr_in, ByRef namelen As Long) As Long
Public Declare Function setsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
Public Declare Function getsockopt Lib "wsock32.dll" (ByVal s As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long

'Communication
Public Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function Send Lib "ws2_32.dll" Alias "send" (ByVal s As Long, ByRef buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long
Public Declare Function recvfrom Lib "ws2_32.dll" (ByVal s As Long, ByRef buf As Byte, ByVal datalen As Long, ByVal Flags As Long, ByRef fromaddr As sockaddr_in, ByRef fromlen As Long) As Long
Public Declare Function sendto Lib "ws2_32.dll" (ByVal s As Long, buf As Any, ByVal lngLen As Long, ByVal Flags As Long, udtTo As sockaddr_in, ByVal tolen As Long) As Long
Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal s As Long, buf As Any, ByVal buflen As Long, ByVal Flags As Long) As Long

'Socket options
Public Declare Function WSAIoctl Lib "ws2_32.dll" (ByVal s As Long, ByVal dwIoControlCode As Long, lpvInBuffer As Any, ByVal cbInBuffer As Long, lpvOutBuffer As Any, ByVal cbOutBuffer As Long, lpcbBytesReturned As Long, lpOverlapped As Long, lpCompletionRoutine As Long) As Long
Public Declare Function vbselect Lib "ws2_32.dll" Alias "select" (ByVal nfds As Long, ByRef readfds As Any, ByRef writefds As Any, ByRef exceptfds As Any, ByRef timeout As Long) As Long
Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long

'Blocking/Async functions
Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long

'Events
Public Declare Function WSACreateEvent Lib "ws2_32.dll" () As Long
Public Declare Function WSACloseEvent Lib "ws2_32.dll" (ByVal hEvent As Long) As Long
Public Declare Function WSAWaitForMultipleEvents Lib "ws2_32.dll" (ByVal cEvents As Long, lphEvents As Long, ByVal fWaitAll As Long, ByVal dwTimeout As Long, ByVal fAlertable As Long) As Long

'Overlapped I/O
Public Declare Function WSAGetOverlappedResult Lib "ws2_32.dll" (ByVal s As Long, lpOverlapped As OVERLAPPED, lpcbTransfer As Long, ByVal fWait As Long, lpdwFlags As Long) As Long

'Errors
Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long


Public saZero               As sockaddr_in
Public WinsockMessage       As Long


'Gets the string vesion of an IP in the form ###.###.###.### given the long version
Public Function GetStrIPFromLong(lngIP As Long) As String
  Dim Bytes(3) As Byte
  
    Call CopyMemory(ByVal VarPtr(Bytes(0)), ByVal VarPtr(lngIP), 4)
    Let GetStrIPFromLong = Bytes(0) & "." & Bytes(1) & "." & Bytes(2) & "." & Bytes(3)
End Function


'Gets the long version of an IP given it in the form ###.###.###.###
Public Function GetLngIPFromStr(strIP As String) As Long
    GetLngIPFromStr = inet_addr(strIP)
End Function


'Returns a list of IP addresses that are available to the local computer.
'Returns in the format IP1;IP2;IP3 etc
'Each IP is of the form ###.###.###.###
Public Function EnumNetworkInterfaces() As String

  Dim lngSocketDescriptor   As Long
  Dim lngBytesReturned      As Long

  Dim Buffer                As INTERFACEINFO
  Dim NumInterfaces         As Integer
  
  Dim i                     As Integer

    lngSocketDescriptor = socket(AF_INET, SOCK_STREAM, 0)
    
    If lngSocketDescriptor = INVALID_SOCKET Then
        EnumNetworkInterfaces = INVALID_SOCKET
        Exit Function
    End If
    
    If WSAIoctl(lngSocketDescriptor, SIO_GET_INTERFACE_LIST, ByVal 0, ByVal 0, Buffer, 1024, lngBytesReturned, ByVal 0, ByVal 0) Then
        EnumNetworkInterfaces = INVALID_SOCKET
        Exit Function
    End If
    
    NumInterfaces = CInt(lngBytesReturned / 76)
   
    For i = 0 To NumInterfaces - 1
        EnumNetworkInterfaces = EnumNetworkInterfaces & GetStrIPFromLong(Buffer.iInfo(i).iiAddress.AddressIn.sin_addr) & ";"
    Next i
    
    closesocket lngSocketDescriptor
    EnumNetworkInterfaces = Left$(EnumNetworkInterfaces, Len(EnumNetworkInterfaces) - 1)

End Function


'Returns a list of IP addresses associated with a hostname.
'Returns in the format IP1;IP2;IP3 etc
'Each IP is of the form ###.###.###.###
Public Function GetIPAddresses(strHostname As String) As String

  Dim lngPtrToHOSTENT As Long
  Dim udtHostent      As HostEnt
  Dim lngPtrToIP      As Long
  Dim arrIpAddress()  As Byte
  Dim strIpAddress    As String
  Dim i               As Integer


    'Call the gethostbyname Winsock API function
    'to get pointer to the HOSTENT structure
    lngPtrToHOSTENT = gethostbyname(strHostname)

    'Check the lngPtrToHOSTENT value
    If lngPtrToHOSTENT <> 0 Then

        'Copy retrieved data to udtHostent structure
        CopyMemory udtHostent, lngPtrToHOSTENT, LenB(udtHostent)

        'Get a pointer to the first address
        CopyMemory lngPtrToIP, udtHostent.h_addr_list, 4

        Do Until lngPtrToIP = 0

            'Prepare the array to receive IP address values
            ReDim arrIpAddress(1 To udtHostent.h_length)

            'move IP address values to the array
            CopyMemory arrIpAddress(1), lngPtrToIP, udtHostent.h_length

            'build string with IP address
            For i = 1 To udtHostent.h_length
                strIpAddress = strIpAddress & arrIpAddress(i) & "."
            Next

            'remove the last dot symbol
            strIpAddress = Left$(strIpAddress, Len(strIpAddress) - 1)

            'Add IP address to the listbox
            GetIPAddresses = GetIPAddresses & strIpAddress & ";"

            'Clear the buffer
            strIpAddress = vbNullString

            'Get pointer to the next address
            udtHostent.h_addr_list = udtHostent.h_addr_list + LenB(udtHostent.h_addr_list)
            CopyMemory lngPtrToIP, udtHostent.h_addr_list, 4

         Loop
         
         GetIPAddresses = Left$(GetIPAddresses, Len(GetIPAddresses) - 1)
    End If

End Function


'Gets the hostname associated with a long address
Public Function GetHostNameByAddr(lngIP As Long) As String
    
  Dim lpHostent  As Long
  Dim udtHostent As HostEnt
    
    'Get the pointer to the HostEnt structure
    lpHostent = gethostbyaddr(lngIP, 4, PF_INET)
    
    'If the pointer is 0 don't continue
    If lpHostent = 0 Then Exit Function
    
    'Copy data into the HostEnt structure
    CopyMemory udtHostent, ByVal lpHostent, LenB(udtHostent)

    'Prepare the buffer to receive a string
    GetHostNameByAddr = String(256, 0)

    'Copy the host name into the strHostName variable
    CopyMemory ByVal GetHostNameByAddr, ByVal udtHostent.h_name, 256

    'Cut received string by first chr(0) character
    GetHostNameByAddr = Left$(GetHostNameByAddr, InStr(1, GetHostNameByAddr, Chr(0)) - 1)
    
End Function


'Returns the long address of strHostname given as either
' - an IP of the form ###.###.###.###
' - A valid host name
Public Function GetHostByNameAlias(ByVal strHostname As String) As Long
  On Error Resume Next
    
  Dim lpHostent   As Long
  Dim udtHostent  As HostEnt
  Dim AddrList    As Long
  Dim retIP       As Long
    
    retIP = inet_addr(strHostname)
    
    If retIP = INADDR_NONE Then
        lpHostent = gethostbyname(strHostname)
        
        If lpHostent <> 0 Then
            CopyMemory udtHostent, ByVal lpHostent, LenB(udtHostent)
            CopyMemory AddrList, ByVal udtHostent.h_addr_list, 4
            CopyMemory retIP, ByVal AddrList, udtHostent.h_length
        Else
            retIP = INADDR_NONE
        End If
    End If
    
    GetHostByNameAlias = retIP

End Function


Public Function GetProtocolName(Protocol As IPProtocol) As String

    Select Case Protocol
             Case IPPROTO_IP:       GetProtocolName = "Dummy for IP"
             Case IPPROTO_ICMP:     GetProtocolName = "Control message protocol"
             Case IPPROTO_IGMP:     GetProtocolName = "Group management protocol"
             Case IPPROTO_GGP:      GetProtocolName = "Gateway^2 (deprecated)"
             Case IPPROTO_TCP:      GetProtocolName = "Tcp"
             Case IPPROTO_PUP:      GetProtocolName = "Pup"
             Case IPPROTO_UDP:      GetProtocolName = "User datagram protocol"
             Case IPPROTO_IDP:      GetProtocolName = "Xns idp"
             Case IPPROTO_ND:       GetProtocolName = "UNOFFICIAL net disk proto"
    End Select

End Function



Public Function GetPortName(Port As ServicePort) As String

    Select Case Port
            Case IPPORT_ECHO: GetPortName = "Echo (" & IPPORT_ECHO & ")"
            Case IPPORT_DISCARD: GetPortName = "Discard (" & IPPORT_DISCARD & ")"
            Case IPPORT_SYSTAT: GetPortName = "Sys Stat (" & IPPORT_SYSTAT & ")"
            Case IPPORT_DAYTIME: GetPortName = "Day Time (" & IPPORT_DAYTIME & ")"
            Case IPPORT_NETSTAT: GetPortName = "NetStat (" & IPPORT_NETSTAT & ")"
            Case IPPORT_FTP: GetPortName = "FTP (" & IPPORT_FTP & ")"
            Case IPPORT_TELNET: GetPortName = "Telnet (" & IPPORT_TELNET & ")"
            Case IPPORT_SMTP: GetPortName = "SMTP (" & IPPORT_SMTP & ")"
            Case IPPORT_TIMESERVER: GetPortName = "Timer Server (" & IPPORT_TIMESERVER & ")"
            Case IPPORT_NAMESERVER: GetPortName = "Name Server (" & IPPORT_NAMESERVER & ")"
            Case IPPORT_WHOIS: GetPortName = "WhoIs (" & IPPORT_WHOIS & ")"
            Case IPPORT_MTP: GetPortName = "MTP (" & IPPORT_MTP & ")"
            Case Else
                GetPortName = CStr(Port)
    End Select

End Function

'Returns whether or not a given string is an IP
'(doesn't verify and IPs existence, just whether or not the string is of IP format ###.###.###.###)
Public Function IsValidIP(strIP As String) As Boolean

  On Error Resume Next

  Dim strIPParts() As String, i As Integer, IPNumber As Byte
    
    'It has to have at least 1 dot in it.
    If InStr(1, strIP, ".") = 0 Then Exit Function
    
    strIPParts = Split(strIP, ".")
    
    'There must be 3 dots to be precise
    If UBound(strIPParts) <> 3 Then Exit Function
        
    For i = 0 To 3
        'The current part must be a number between 0 and 255
        IPNumber = CByte(strIPParts(i))
        If Err Then Exit Function
    Next i
    
    IsValidIP = True

End Function
