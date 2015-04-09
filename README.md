#Note: This repo is now deprecated and no longer supported. 

# semantics3-classicasp

semantics3-classicasp is a Classic ASP VBScript code sample for accessing the Semantics3 Products API, which provides structured information, including pricing histories, for a large number of products.

Note: This is just a code example on making an oauth request to the Semantics3 library. and is not a library.

See https://www.semantics3.com for more information.

Quickstart guide: https://www.semantics3.com/quickstart
API documentation can be found at https://www.semantics3.com/docs/

## Download

semantics3-classicasp can be cloned from this link:
```bash
$ git clone git@github.com:Semantics3/semantics3-classicasp.git
```

## Requirements

* IIS (or IIS Express)
* VBScriptOAuth (http://scottdesapio.com/VBScriptOAuth/oAuthASPExample.zip, project homepage http://scottdesapio.com/VBScriptOAuth/)

## Getting Started

In order to use the client, you must have both an API key and an API secret. To obtain your key and secret, you need to first create an account at
https://www.semantics3.com/
You can access your API access credentials from the user dashboard at https://www.semantics3.com/dashboard/applications

### Setup Work

This code sample uses Classic ASP VBScript OAuth library by Scott DeSapio (http://scottdesapio.com/VBScriptOAuth/) to make the OAuth authentication with Semantics3 Products API. The OAuth library can be downloaded from http://scottdesapio.com/VBScriptOAuth/oAuthASPExample.zip or from this github repo.
Then it constructs the JSON query, queries the products endpoint and prints the API reponse to the web page.


```vbscript
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
```

## Contributing

Use GitHub's standard fork/commit/pull-request cycle.  If you have any questions, email <support@semantics3.com>.

## Author

* Srinivasan Kidambi <srinivas@semantics3.com>

## Copyright

Copyright (c) 2013 Semantics3 Inc.

## License

    The "MIT" License
    
    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:
    
    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.
    
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS
    OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL
    THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
    DEALINGS IN THE SOFTWARE.

