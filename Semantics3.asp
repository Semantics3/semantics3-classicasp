<!--#include file="oauth/cLibOAuth.asp"-->

<%
    
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 1. Instantiate an instance of the OAuth object.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim objOAuth : Set objOAuth = New cLibOAuth
objOAuth.ConsumerKey = "API_KEY"
objOAuth.ConsumerSecret = "API_SECRET"
objOAuth.EndPoint = "https://api.semantics3.com/v1/products"
objOAuth.RequestMethod = "GET"
objOAuth.UserAgent = "OAuth gem v1.0"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 2. Add proprietary request parameters.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Dim paramString : paramString = "{""cat_id"":13658,""brand"":""Toshiba"",""model"":""Satellite""}"
paramString = Server.URLEncode(paramString)

objOAuth.Parameters.Add "q", paramString

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 3. Make the call.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
objOAuth.Send()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 4. Evaluate the response.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Dim strResponse : strResponse = objOAuth.Get_ResponseValue("RESPONSE_PARAM_NAME")

Dim strResponseText : strResponseText = objOAuth.ResponseText
Dim strErrorCode : strErrorCode = objOAuth.ErrorCode

If Not IsNull(strErrorCode) Then
    Response.Status = RESPONSE_STATUS_500
    Response.Write strErrorCode
Else
    Response.ContentType = "text/html"
    Response.CharSet = "utf-8"
    Response.Write strResponseText
End If

%>