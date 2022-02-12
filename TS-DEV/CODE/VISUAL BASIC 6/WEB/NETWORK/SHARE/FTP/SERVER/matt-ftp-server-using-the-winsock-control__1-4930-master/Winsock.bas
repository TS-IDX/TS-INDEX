Attribute VB_Name = "Winsock"
Option Explicit

Global Const AF_UNSPEC = 0             '  /* unspecified */
Global Const AF_UNIX = 1               '  /* local to host (pipes, portals) */
Global Const AF_INET = 2               '  /* internetwork: UDP, TCP, etc. */
Global Const AF_IMPLINK = 3            '  /* arpanet imp addresses */
Global Const AF_PUP = 4                '  /* pup protocols: e.g. BSP */
Global Const AF_CHAOS = 5              '  /* mit CHAOS protocols */
Global Const AF_IPX = 6                '  /* IPX and SPX */
Global Const AF_NS = 6                 '  /* XEROX NS protocols */
Global Const AF_ISO = 7                '  /* ISO protocols */
Global Const AF_OSI = AF_ISO           '  /* OSI is ISO */
Global Const AF_ECMA = 8               '  /* european computer manufacturers */
Global Const AF_DATAKIT = 9            '  /* datakit protocols */
Global Const AF_CCITT = 10             '  /* CCITT protocols, X.25 etc */
Global Const AF_SNA = 11               '  /* IBM SNA */
Global Const AF_DECnet = 12            '  /* DECnet */
Global Const AF_DLI = 13               '  /* Direct data link interface */
Global Const AF_LAT = 14               '  /* LAT */
Global Const AF_HYLINK = 15            '  /* NSC Hyperchannel */
Global Const AF_APPLETALK = 16         '  /* AppleTalk */
Global Const AF_NETBIOS = 17           '  /* NetBios-style addresses */

Global Const FD_READ = &H1
Global Const FD_WRITE = &H2
Global Const FD_OOB = &H4
Global Const FD_ACCEPT = &H8
Global Const FD_CONNECT = &H10
Global Const FD_CLOSE = &H20
Global Const FD_SETSIZE% = 64

Public Const SOL_SOCKET = &HFFFF
Public Const SO_LINGER = &H80

Global Const INVALID_SOCKET = -1
Global Const SOCKET_ERROR = -1

Global Const BAD_SOCKET = -1
Global Const UNRESOLVED_HOST = -2
Global Const UNABLE_TO_BIND = -3
Global Const UNABLE_TO_CONNECT = -4

 
Global Const WIN_SOCKET_MSG = 2000
Public Const MAX_WSADescription = 257
Public Const MAX_WSASYSStatus = 129

Public Const WS_VERSION_REQD As Integer = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD / &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const IP_OPTIONS = 1
Public Const MIN_SOCKETS_REQD = 0

'--- additional declarations
'Types
Global Const SOCK_STREAM = 1
Global Const SOCK_DGRAM = 2
Global Const SOCK_RAW = 3
Global Const SOCK_RDM = 4
Global Const SOCK_SEQPACKET = 5

'Protocol families, same as address families for now
Global Const PF_UNSPEC = 0
Global Const PF_UNIX = 1
Global Const PF_INET = 2
Global Const PF_IMPLINK = 3
Global Const PF_PUP = 4
Global Const PF_CHAOS = 5
Global Const PF_IPX = 6
Global Const PF_NS = 6
Global Const PF_ISO = 7
Global Const PF_OSI = AF_ISO
Global Const PF_ECMA = 8
Global Const PF_DATAKIT = 9
Global Const PF_CCITT = 10
Global Const PF_SNA = 11
Global Const PF_DECnet = 12
Global Const PF_DLI = 13
Global Const PF_LAT = 14
Global Const PF_HYLINK = 15
Global Const PF_APPLETALK = 16
Global Const PF_NETBIOS = 17

Public Const MAXGETHOSTSTRUCT = 1024

Public Const IPPROTO_TCP = 6
Public Const IPPROTO_UDP = 17

Public Const INADDR_NONE = &HFFFF
Public Const INADDR_ANY = &H0

' Windows Sockets definitions of regular Microsoft C error constants
Global Const WSAEINTR = 10004
Global Const WSAEBADF = 10009
Global Const WSAEACCES = 10013
Global Const WSAEFAULT = 10014
Global Const WSAEINVAL = 10022
Global Const WSAEMFILE = 10024
' Windows Sockets definitions of regular Berkeley error constants
Global Const WSAEWOULDBLOCK = 10035
Global Const WSAEINPROGRESS = 10036
Global Const WSAEALREADY = 10037
Global Const WSAENOTSOCK = 10038
Global Const WSAEDESTADDRREQ = 10039
Global Const WSAEMSGSIZE = 10040
Global Const WSAEPROTOTYPE = 10041
Global Const WSAENOPROTOOPT = 10042
Global Const WSAEPROTONOSUPPORT = 10043
Global Const WSAESOCKTNOSUPPORT = 10044
Global Const WSAEOPNOTSUPP = 10045
Global Const WSAEPFNOSUPPORT = 10046
Global Const WSAEAFNOSUPPORT = 10047
Global Const WSAEADDRINUSE = 10048
Global Const WSAEADDRNOTAVAIL = 10049
Global Const WSAENETDOWN = 10050
Global Const WSAENETUNREACH = 10051
Global Const WSAENETRESET = 10052
Global Const WSAECONNABORTED = 10053
Global Const WSAECONNRESET = 10054
Global Const WSAENOBUFS = 10055
Global Const WSAEISCONN = 10056
Global Const WSAENOTCONN = 10057
Global Const WSAESHUTDOWN = 10058
Global Const WSAETOOMANYREFS = 10059
Global Const WSAETIMEDOUT = 10060
Global Const WSAECONNREFUSED = 10061
Global Const WSAELOOP = 10062
Global Const WSAENAMETOOLONG = 10063
Global Const WSAEHOSTDOWN = 10064
Global Const WSAEHOSTUNREACH = 10065
Global Const WSAENOTEMPTY = 10066
Global Const WSAEPROCLIM = 10067
Global Const WSAEUSERS = 10068
Global Const WSAEDQUOT = 10069
Global Const WSAESTALE = 10070
Global Const WSAEREMOTE = 10071
' Extended Windows Sockets error constant definitions
Global Const WSASYSNOTREADY = 10091
Global Const WSAVERNOTSUPPORTED = 10092
Global Const WSANOTINITIALISED = 10093
Global Const WSAHOST_NOT_FOUND = 11001
Global Const WSATRY_AGAIN = 11002
Global Const WSANO_RECOVERY = 11003
Global Const WSANO_DATA = 11004
Global Const WSANO_ADDRESS = 11004

Type hostent
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
End Type

Public hostent As hostent

Type WSAdata
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * MAX_WSADescription '(0 To 255) As Byte
    szSystemStatus As String * MAX_WSASYSStatus  '(0 To 128) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Public WSAdata As WSAdata

Type Inet_Address     ' IP Address in Network Order
    Byte4 As Byte     '
    Byte3 As Byte     '
    Byte2 As Byte     '
    Byte1 As Byte     '
End Type

Public IPLong As Inet_Address


'socket address
Type SockAddr
    sin_family As Integer   ' Address family
    sin_port As Integer     ' Port Number in Network Order
    sin_addr As Long        ' IP Address as Long
    sin_zero As String * 8  '(8) As Byte             ' Padding
End Type

Public SockAddr As SockAddr

Public Const SockAddr_Size = 16

Type hostent_async
    h_name As Long
    h_aliases As Long
    h_addrtype As Integer
    h_length As Integer
    h_addr_list As Long
    h_asyncbuffer(MAXGETHOSTSTRUCT) As Byte
End Type

Public hostent_async As hostent_async

Type fd_set
  fd_count As Integer          '' how many are in the set
  fd_array(FD_SETSIZE) As Long '' array of SOCKET handles (64)
End Type

Public fd_set As fd_set

Type timeval
    tv_sec As Long
    tv_usec As Long
End Type

Public timeval As timeval

Type LingerType
    l_onoff As Integer
    l_linger As Integer
End Type

'---SOCKET FUNCTIONS
    Public Declare Function accept Lib "wsock32.dll" (ByVal S As Long, addr As SockAddr, addrlen As Long) As Long
    Public Declare Function bind Lib "wsock32.dll" (ByVal S As Long, addr As SockAddr, ByVal namelen As Long) As Long
    Public Declare Function closesocket Lib "wsock32.dll" (ByVal S As Long) As Long
    Public Declare Function connect Lib "wsock32.dll" (ByVal S As Long, addr As SockAddr, ByVal namelen As Long) As Long
    Public Declare Function ioctlsocket Lib "wsock32.dll" (ByVal S As Long, ByVal cmd As Long, argp As Long) As Long
    Public Declare Function getpeername Lib "wsock32.dll" (ByVal S As Long, sName As SockAddr, namelen As Long) As Long
    Public Declare Function getsockname Lib "wsock32.dll" (ByVal S As Long, sName As SockAddr, namelen As Long) As Long
    Public Declare Function getsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, optlen As Long) As Long
    Public Declare Function htonl Lib "wsock32.dll" (ByVal hostlong As Long) As Long
    Public Declare Function htons Lib "wsock32.dll" (ByVal hostshort As Long) As Integer
    Public Declare Function inet_addr Lib "wsock32.dll" (ByVal CP As String) As Long
    Public Declare Function inet_ntoa Lib "wsock32.dll" (ByVal inn As Long) As Long
    Public Declare Function listen Lib "wsock32.dll" (ByVal S As Long, ByVal backlog As Long) As Long
    Public Declare Function ntohl Lib "wsock32.dll" (ByVal netlong As Long) As Long
    Public Declare Function ntohs Lib "wsock32.dll" (ByVal netshort As Long) As Integer
    Public Declare Function recv Lib "wsock32.dll" (ByVal S As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function recvfrom Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, from As SockAddr, fromlen As Long) As Long
    Public Declare Function ws_select Lib "wsock32.dll" Alias "select" (ByVal nfds As Long, readfds As fd_set, writefds As fd_set, exceptfds As fd_set, timeout As timeval) As Long
    Public Declare Function send Lib "wsock32.dll" (ByVal S As Long, ByVal buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long
    Public Declare Function sendto Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long, to_addr As SockAddr, ByVal tolen As Long) As Long
    Public Declare Function setsockopt Lib "wsock32.dll" (ByVal S As Long, ByVal Level As Long, ByVal optname As Long, optval As Any, ByVal optlen As Long) As Long
    Public Declare Function ShutDown Lib "wsock32.dll" Alias "shutdown" (ByVal S As Long, ByVal how As Long) As Long
    Public Declare Function Socket Lib "wsock32.dll" Alias "socket" (ByVal af As Long, ByVal s_type As Long, ByVal protocol As Long) As Long
'---DATABASE FUNCTIONS
    Public Declare Function gethostbyaddr Lib "wsock32.dll" (addr As Long, ByVal addr_len As Long, ByVal addr_type As Long) As Long
    Public Declare Function gethostbyname Lib "wsock32.dll" (ByVal host_name As String) As Long
    Public Declare Function gethostname Lib "wsock32.dll" (ByVal host_name As String, ByVal namelen As Long) As Long
    Public Declare Function getservbyport Lib "wsock32.dll" (ByVal Port As Long, ByVal proto As String) As Long
    Public Declare Function getservbyname Lib "wsock32.dll" (ByVal serv_name As String, ByVal proto As String) As Long
    Public Declare Function getprotobynumber Lib "wsock32.dll" (ByVal proto As Long) As Long
    Public Declare Function getprotobyname Lib "wsock32.dll" (ByVal proto_name As String) As Long
'---WINDOWS EXTENSIONS
    Public Declare Function WSAStartup Lib "wsock32.dll" (ByVal wVR As Long, lpWSAD As WSAdata) As Long
    Public Declare Function WSACleanup Lib "wsock32.dll" () As Long
    Public Declare Function WSASetLastError Lib "wsock32.dll" (ByVal iError As Long) As Long
    Public Declare Function WSAGetLastError Lib "wsock32.dll" () As Long
    Public Declare Function WSAIsBlocking Lib "wsock32.dll" () As Long
    Public Declare Function WSAUnhookBlockingHook Lib "wsock32.dll" () As Long
    Public Declare Function WSASetBlockingHook Lib "wsock32.dll" (ByVal lpBlockFunc As Long) As Long
    Public Declare Function WSACancelBlockingCall Lib "wsock32.dll" () As Long
    Public Declare Function WSAAsyncGetServByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal serv_name As String, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetServByPort Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Port As Long, ByVal proto As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal proto_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetProtoByNumber Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal Number As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByName Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal host_name As String, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSAAsyncGetHostByAddr Lib "wsock32.dll" (ByVal hWnd As Long, ByVal wMsg As Long, addr As Long, ByVal addr_len As Long, ByVal addr_type As Long, buf As Any, ByVal buflen As Long) As Long
    Public Declare Function WSACancelAsyncRequest Lib "wsock32.dll" (ByVal hAsyncTaskHandle As Long) As Long
    Public Declare Function WSAAsyncSelect Lib "wsock32.dll" (ByVal S As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
    Public Declare Function WSARecvEx Lib "wsock32.dll" (ByVal S As Long, buf As Any, ByVal buflen As Long, ByVal FLAGS As Long) As Long

