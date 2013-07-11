VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OAuth2GoogleAutheticator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAuthenticator


Option Explicit


' --------------------------------------------- '
' Properties
' --------------------------------------------- '


Public ClientId As String
Public ClientSecret As String
Public Username As String
Public Password As String
Public TokenUrl As String
Public TokenKey As String
Public Token As String
Public AuthUrl As String
Public AuthCode As String
Public CacheToken As Boolean


Private ieComplete As Boolean


Private WithEvents oIExplorer As InternetExplorer
Attribute oIExplorer.VB_VarHelpID = -1


' ============================================= '
' Public Methods
' ============================================= '


''
' Helper for setting required parameters at once
'
' @param {String} ClientId
' @param {String} ClientSecret
' --------------------------------------------- '


Public Sub Setup(ClientId As String, ClientSecret As String)
    Me.ClientId = ClientId
    Me.ClientSecret = ClientSecret
End Sub


''
' Setup token url that is used to request token
'
' @param (String) AuthUrl
' @param {String} TokenUrl
' @param {String} [TokenKey="access_token"] The key used to retrieve the token from the response
' --------------------------------------------- '


Public Sub SetupTokenUrl(AuthUrl As String, TokenUrl As String, Optional TokenKey As String = "access_token")
    Me.TokenUrl = TokenUrl
    Me.TokenKey = TokenKey
    Me.AuthUrl = AuthUrl
End Sub


Private Sub IAuthenticator_BeforeExecute(Request As RestRequest)
    On Error GoTo ErrorHandling
    If (Me.Token = "" Or Not Me.CacheToken) And (Me.TokenUrl <> "" And Me.TokenKey <> "") Then
        
        Set oIExplorer = New InternetExplorer
        With oIExplorer
            .Navigate CreateAuthRequest()
            .AddressBar = False
            .MenuBar = False
            .Resizable = False
            .Visible = True
        End With
        
        'Wait for userInteraction
        Do: DoEvents: Loop Until ieComplete
        
        'Do we have an Authentication Code?
        If Len(Me.AuthCode) = 0 Then
            Err.Raise vbObjectError + 2, _
                Description:="User cancelled Authentication"
        End If
        
        'Now Get a new Token
        Dim http As MSXML2.ServerXMLHTTP
        Set http = New MSXML2.ServerXMLHTTP


        Call http.Open("POST", Me.TokenUrl, False)
        http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
        Call http.send(CreateTokenRequest())


        If http.Status <> 200 Then
            ' Error getting OAuth2 token
            Err.Raise vbObjectError + http.Status, _
                Description:="Failed to retrieve OAuth2 Token - " & http.Status & ": " & http.responseText
        End If
        
        Dim Response As Dictionary
        Set Response = RestHelpers.ParseJSON(http.responseText)
        If Not Response Is Nothing Then
            If Response.Exists(Me.TokenKey) Then
                Me.Token = Response(Me.TokenKey)
            End If
            
            ' (Salesforce specific, but shouldn't affect any other OAuth2 clients)
            If Response.Exists("instance_url") Then
                Request.BaseUrl = Response("instance_url")
            End If
        Else
            Err.Raise vbObjectError + 2, _
                Description:="Failed to read OAuth2 Token"
        End If
    End If
    Call Request.AddHeader("Authorization", CreateHeader())
    
ErrorHandling:
    
    If Not http Is Nothing Then Set http = Nothing
    If Not Response Is Nothing Then Set Response = Nothing
    
    If Err.Number <> 0 Then
        ' Rethrow error
        Err.Raise Err.Number, Description:=Err.Description
    End If
End Sub


Private Sub IAuthenticator_HttpOpen(http As MSXML2.IXMLHTTPRequest, Request As RestRequest, BaseUrl As String, Optional useAsync As Boolean = False)
    ' Perform standard http open
    Call http.Open(Request.MethodName(), Request.FullUrl(BaseUrl), useAsync)
End Sub


' ============================================= '
' Private Methods
' ============================================= '


Private Function CreateHeader() As String
    ' Create standard OAuth2 header
    CreateHeader = "Bearer " & Me.Token
End Function
Private Function CreateAuthRequest() As String
    ' Generate initial Authentication Request
    ' Using installed application flow: https://developers.google.com/accounts/docs/OAuth2InstalledApp
    CreateAuthRequest = Me.AuthUrl
    If InStr(1, CreateAuthRequest, "?") < 1 Then: CreateAuthRequest = CreateAuthRequest & "?"
    CreateAuthRequest = CreateAuthRequest & "response_type=code"
    CreateAuthRequest = CreateAuthRequest & "&client_id=" & Me.ClientId
    CreateAuthRequest = CreateAuthRequest & "&&redirect_uri=urn:ietf:wg:oauth:2.0:oob"
    CreateAuthRequest = CreateAuthRequest & "&scope=openid"
End Function




Private Function CreateTokenRequest() As String


    CreateTokenRequest = "code=" & Me.AuthCode
    CreateTokenRequest = CreateTokenRequest & "&client_id=" & Me.ClientId
    CreateTokenRequest = CreateTokenRequest & "&client_secret=" & Me.ClientSecret
    CreateTokenRequest = CreateTokenRequest & "&redirect_uri=urn:ietf:wg:oauth:2.0:oob" 'More Likely than not to be static
    CreateTokenRequest = CreateTokenRequest & "&grant_type=authorization_code" 'More Likely than not to be static


End Function
'Break Loop on user Quit of IE
Private Sub oIExplorer_OnQuit()
    ieComplete = True
End Sub


Private Sub oIExplorer_TitleChange(ByVal Text As String)
    If InStr(1, Text, "Success") > 0 Then
        Me.AuthCode = oIExplorer.Document.getElementbyid("code").Value
        oIExplorer.Quit
    ElseIf InStr(1, Text, "Denied") > 0 Then
        oIExplorer.Quit
    End If
End Sub



