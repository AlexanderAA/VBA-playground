Function getObjectParameter(object_id) As String
    ' XXX: Add validation for object_id
    
    ' === Settings
    UserName = ""
    Password = ""
    SBUrl = "https://domain.tld/"
    ProxyUrl = "127.0.0.1:3128"
    CandStatusUrl = SBUrl + "excelapi/cs/"
    Dim csrftoken As String
    Dim sessionid As String

    ' === First request - get login page with csrf token and session id
    Dim r1 As Object
    Set r1 = CreateObject("WinHttp.WinHttpRequest.5.1")
    r1.Open "GET", SBUrl, True
    r1.SetProxy 2, ProxyUrl, ""
    r1.Send
    r1.WaitForResponse
    csrftoken = r1.GetResponseHeader("csrftoken")
    sessionid = r1.GetResponseHeader("sessionid")
    
    ' === Second request - Log in
    ' Submit our username and password with csrf token and session id from the first step
    r1.Open "POST", SBUrl, True
    ' Very important option is below. Unfortunately, I was too lazy to find out its name.
    ' Without Option(6) = False the whole thing stops working with error "80072ef1" as soon as HTTP 302 redirect is discovered.
    r1.Option(6) = False ' Do not handle redirects.. Because you can't handle them.
    r1.SetProxy 2, ProxyUrl, ""
    'r1.SetRequestHeader "Cookie", "csrftoken=" + csrftoken
    r1.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    r1.Send ("csrfmiddlewaretoken=" + csrftoken + "&username=" + UserName + "&password=" + Password + "&this_is_the_login_form=1&next=%2F")
    r1.WaitForResponse
    csrftoken = r1.GetResponseHeader("csrftoken")
    sessionid = r1.GetResponseHeader
    
    ' === Third request - ask for object property
    r1.Open "GET", CandStatusUrl + object_id + "/", False
    r1.SetRequestHeader "Cookie", "sessionid=" + sessionid + ";csrftoken=" + csrftoken
    r1.Send
    ' Fetch result
    'hdr = r1.GetResponseHeader("sessionid")
    getObjectParameter = r1.ResponseText
End Function
