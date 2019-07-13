VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Net"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "Net"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const BUFFER_LENGTH                 As Long = 1024
Private Const FIELDS_BUFFER_LENGTH          As Long = 10
Private Const FIELDS_NAME_INDEX             As Long = 0
Private Const FIELDS_VALUE_INDEX            As Long = 1
Private Const INTERNET_OPEN_TYPE_DIRECT     As Long = 1
Private Const INTERNET_SERVICE_HTTP         As Long = 3
Private Const INTERNET_FLAG_PRAGMA_NOCACHE  As Long = &H100
Private Const INTERNET_FLAG_KEEP_CONNECTION As Long = &H400000
Private Const INTERNET_FLAG_SECURE          As Long = &H800000
Private Const INTERNET_FLAG_FROM_CACHE      As Long = &H1000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE  As Long = &H4000000
Private Const INTERNET_FLAG_RELOAD          As Long = &H80000000
Private Declare Function apiInternetCloseHandle Lib "wininet" Alias "InternetCloseHandle" (ByVal hInet As Long) As Boolean
Private Declare Function apiHttpOpenRequest Lib "wininet" Alias "HttpOpenRequestA" (ByVal hHttpSession As Long, ByVal Verb As String, ByVal ObjectName As String, ByVal Version As String, ByVal Referer As String, ByVal AcceptTypes As Long, ByVal Flags As Long, Context As Long) As Long
Private Declare Function apiHttpSendRequest Lib "wininet" Alias "HttpSendRequestA" (ByVal hHttpRequest As Long, ByVal Headers As String, ByVal HeadersLength As Long, ByVal sOptional As String, ByVal OptionalLength As Long) As Boolean
Private Declare Function apiInternetConnect Lib "wininet" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal ServerName As String, ByVal ServerPort As Integer, ByVal UserName As String, ByVal Password As String, ByVal Service As Long, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function apiInternetOpen Lib "wininet" Alias "InternetOpenA" (ByVal Agent As String, ByVal AccessType As Long, ByVal ProxyName As String, ByVal ProxyBypass As String, ByVal Flags As Long) As Long
Private Declare Function apiInternetReadFile Lib "wininet" Alias "InternetReadFile" (ByVal hConnect As Long, ByVal buffer As String, ByVal NumberOfBytesToRead As Long, NumberOfBytesRead As Long) As Boolean
Private hHTTP           As Long
Private hConnection     As Long
Private DontEncode(255) As Boolean
Private FieldCount      As Long
Private mFields()       As String
Public Enum ePort
    INTERNET_DEFAULT_HTTP_PORT = 80
    INTERNET_DEFAULT_HTTPS_PORT = 443
End Enum
Private Sub Class_Initialize()
    Dim l As Long
    ReDim mFields(1, FIELDS_BUFFER_LENGTH)
    For l = Asc("0") To Asc("9")
        DontEncode(l) = True
    Next
    For l = Asc("a") To Asc("z")
        DontEncode(l) = True
    Next
    For l = Asc("A") To Asc("Z")
        DontEncode(l) = True
    Next
End Sub
Private Sub Class_Terminate()
    Erase mFields
End Sub
Friend Property Let Fields(ByVal Name As String, Value As String)
    mFields(FIELDS_VALUE_INDEX, GetFieldIndex(Name, True)) = Value
End Property
Friend Property Get Fields(ByVal Name As String) As String
    Dim l As Long
    l = GetFieldIndex(Name, False)
    If l > -1 Then Fields = mFields(FIELDS_VALUE_INDEX, l)
End Property
Friend Function OpenHTTP(ByVal Server As String, Optional Port As ePort = INTERNET_DEFAULT_HTTP_PORT, Optional UserName As String, Optional Password As String) As Boolean
    CloseHTTP
    hHTTP = apiInternetOpen("HTTP Client", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    If hHTTP = 0 Then Exit Function
    hConnection = apiInternetConnect(hHTTP, Server, INTERNET_DEFAULT_HTTP_PORT, UserName, Password, INTERNET_SERVICE_HTTP, 0, 0)
    If hConnection <> 0 Then
        OpenHTTP = True
    Else
        apiInternetCloseHandle hHTTP
        hHTTP = 0
    End If
End Function
Friend Sub CloseHTTP()
    If hConnection <> 0 Then apiInternetCloseHandle hConnection
    hConnection = 0
    If hHTTP <> 0 Then apiInternetCloseHandle hHTTP
    hHTTP = 0
End Sub
Friend Function SendRequest(ByVal File As String, Optional Method As String = "GET", Optional Referer As String, Optional Reload As Boolean = True) As String
    Dim hRequest As Long
    Dim r        As Boolean
    Dim buffer   As String
    Dim Header   As String
    Dim Request  As String
    Dim POSTData As String
    Dim Response As String
    Dim Read     As Long
    Dim Flags    As Long
    Method = UCase(Method)
    Request = BuildRequest
    buffer = Space(BUFFER_LENGTH)
    If Len(Request) > 0 Then
        If Method = "POST" Then
            Header = "Content-Type: application/x-www-form-urlencoded"
            POSTData = Request
        Else
            File = File & "?" & Request
        End If
    End If
    If Reload = True Then Flags = Flags Or INTERNET_FLAG_PRAGMA_NOCACHE Or INTERNET_FLAG_RELOAD
    hRequest = apiHttpOpenRequest(hConnection, Method, File, "HTTP/1.1", "", 0, Flags, 0)
    If hRequest <> 0 Then
        If apiHttpSendRequest(hRequest, Header, Len(Header), POSTData, Len(POSTData)) Then
            r = apiInternetReadFile(hRequest, buffer, BUFFER_LENGTH, Read)
            While r And (Read <> 0)
                Response = Response & Left(buffer, Read)
                r = apiInternetReadFile(hRequest, buffer, BUFFER_LENGTH, Read)
            Wend
        End If
        apiInternetCloseHandle hRequest
    End If
    SendRequest = Response
End Function
Private Function GetFieldIndex(ByVal Name As String, Optional Add As Boolean) As Long
    Dim l As Long
    For l = 0 To FieldCount - 1
        If StrComp(Name, mFields(FIELDS_NAME_INDEX, l), vbTextCompare) = 0 Then
            GetFieldIndex = l
            Exit Function
        End If
    Next
    If Add Then
        If FieldCount = UBound(mFields, 2) Then
            ReDim Preserve mFields(1, UBound(mFields, 2) + FIELDS_BUFFER_LENGTH)
        End If
        mFields(FIELDS_NAME_INDEX, FieldCount) = Name
        GetFieldIndex = FieldCount
        FieldCount = FieldCount + 1
    Else
        GetFieldIndex = -1
    End If
End Function
Private Function BuildRequest() As String
    Dim l As Long
    Dim s As String
    For l = 0 To FieldCount - 1
        s = s & URLEncode(mFields(FIELDS_NAME_INDEX, l)) & "=" & URLEncode(mFields(FIELDS_VALUE_INDEX, l)) & "&"
    Next
    If Len(s) > 0 Then BuildRequest = Left(s, Len(s) - 1)
End Function
Friend Function URLEncode(ByVal Data As String) As String
    Dim l   As Long
    Dim b() As Byte
    Dim s   As String
    Dim C   As String
    b = Data
    For l = 0 To UBound(b) Step 2
        If DontEncode(b(l)) = True Then
            s = s & Chr(b(l))
        Else
            C = Hex(b(l))
            While Len(C) < 2
                C = "0" & C
            Wend
            s = s & "%" & C
        End If
    Next
    URLEncode = s
    
    
End Function

