# 示例代码

```vb
Sub    TestHttpClient()
        On Error GoTo EH
        With New cHttpClient
            'Make json body
            With .RequestDataJson
                With .NewItem("consignee")
                    With .NewItem("address")
                        .Item("city") = "Paderborn"
                        .Item("countryCode") = "DE"
                        .Item("postCode") = "33100"
                    End With
                End With
                With .NewItems("lines")
                    With .NewItem()
                        .Item("content") = "furniture"
                        .Item("unitHeight") = 120
                        .Item("unitLength") = 120
                        .Item("unitPackageType") = "EP"
                        .Item("unitQuantity") = 1
                        .Item("unitWeight") = 200
                        .Item("unitWidth") = 80
                    End With
                End With
                .Item("product") = "STANDARD"
                With .NewItem("consignee")
                    With .NewItem("address")
                        .Item("city") = "Berlin"
                        .Item("countryCode") = "DE"
                        .Item("postCode") = "10115"
                    End With
                End With
                .Item("pickupOn") = Format$(Now, "yyyy-MM-dd")
                'Check json data
                Debug.Print "Request Data: ",.Encode(, 2, True)
            End With
            'You can open debug info object form build in.
            .DebugStart = True
            'Add any header
            .RequestHeaders.Item("X-API-KEY") = "YOUR API KEY HERE"
            'You can set Content-Type like this:
            .SetRequestContentType(EnumRequestContentType.JsonString)
            'Post
            .SendPost("https://vb6.pro/echo/?hello=woeoio")
            'Server retun json data
            With .ReturnJson()
                MsgBox(.Root("server")("time"))
                'Show server reopnse, you can use .ReturnText()
                Debug.Print "Server Response: ", .Encode(, 2, True)
            End With
            QH:
            'Show deubg info
            Debug.Print "Debug Info: ", .DebugInfo.Encode(, 2, True)
            Exit Sub
            EH:
            'Do some thing...
            MsgBox(Err.Description)
            Resume QH
        End With
    End Sub
```

## 服务器返回

```log
17:26:14.470 TestHttpClient()
17:26:14.511 Request Data:               {
  "consignee": {
    "address": {
      "city": "Berlin",
      "countryCode": "DE",
      "postCode": "10115"
    }
  },
  "lines": [
    {
      "content": "furniture",
      "unitHeight": 120,
      "unitLength": 120,
      "unitPackageType": "EP",
      "unitQuantity": 1,
      "unitWeight": 200,
      "unitWidth": 80
    }
  ],
  "product": "STANDARD",
  "pickupOn": "2026-02-08"
}
17:26:16.494 Server Response:            {
  "server": {
    "time": "2026-02-08 17:26:14 CST",
    "timestamp": 1770542774,
    "timezone": "PRC",
    "software": "nginx/1.22.1",
    "name": "www.vb6.pro",
    "address": "172.23.151.44",
    "port": "443",
    "protocol": "HTTP/1.1",
    "document_root": "/www/wwwroot/vb6.pro",
    "php_version": "8.2.28",
    "sapi": "fpm-fcgi",
    "os": "Linux"
  },
  "request": {
    "method": "POST",
    "uri": "/echo/?hello=woeoio",
    "url": "https://",
    "query_string": "hello=woeoio",
    "path_info": "",
    "script_name": "/echo/index.php",
    "filename": "/www/wwwroot/vb6.pro/echo/index.php",
    "time": "2026-02-08 17:26:14 CST",
    "timestamp": 1770542774,
    "time_float": 1770542774.17496,
    "http_version": "HTTP/1.1",
    "is_ajax": false,
    "is_secure": true,
    "full_url": "https://vb6.pro/echo/?hello=woeoio",
    "base_url": "https://vb6.pro/echo"
  },
  "client": {
    "ip": "61.146.189.212",
    "port": "21702",
    "user_agent": "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)",
    "referer": "",
    "language": "",
    "encoding": "",
    "charset": "",
    "connection": "Keep-Alive",
    "host": "vb6.pro"
  },
  "headers": {
    "Host": "vb6.pro",
    "Content-Length": "266",
    "Content-Lenght": "266",
    "X-Api-Key": "YOUR API KEY HERE",
    "User-Agent": "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)",
    "Accept": "*/*",
    "Content-Type": "application/json",
    "Connection": "Keep-Alive"
  },
  "get": {
    "hello": "woeoio"
  },
  "post": {
    "consignee": {
      "address": {
        "city": "Berlin",
        "countryCode": "DE",
        "postCode": "10115"
      }
    },
    "lines": [
      {
        "content": "furniture",
        "unitHeight": 120,
        "unitLength": 120,
        "unitPackageType": "EP",
        "unitQuantity": 1,
        "unitWeight": 200,
        "unitWidth": 80
      }
    ],
    "product": "STANDARD",
    "pickupOn": "2026-02-08"
  },
  "files": [
  ],
  "cookies": [
  ],
  "session": [
  ],
  "body": "{\"consignee\":{\"address\":{\"city\":\"Berlin\",\"countryCode\":\"DE\",\"postCode\":\"10115\"}},\"lines\":[{\"content\":\"furniture\",\"unitHeight\":120,\"unitLength\":120,\"unitPackageType\":\"EP\",\"unitQuantity\":1,\"unitWeight\":200,\"unitWidth\":80}],\"product\":\"STANDARD\",\"pickupOn\":\"2026-02-08\"}",
  "parsed_body": {
    "json": {
      "consignee": {
        "address": {
          "city": "Berlin",
          "countryCode": "DE",
          "postCode": "10115"
        }
      },
      "lines": [
        {
          "content": "furniture",
          "unitHeight": 120,
          "unitLength": 120,
          "unitPackageType": "EP",
          "unitQuantity": 1,
          "unitWeight": 200,
          "unitWidth": 80
        }
      ],
      "product": "STANDARD",
      "pickupOn": "2026-02-08"
    }
  },
  "content": {
    "type": "application/json",
    "length": "266",
    "md5": "fbf19d79627687cd67ddd6eaa2971f43",
    "encoding": ""
  },
  "auth": {
    "type": "",
    "user": ""
  },
  "ssl": {
    "protocol": "Unknown",
    "cipher": "Unknown",
    "version": "Unknown",
    "verify_client": "NONE"
  },
  "phpinfo": [
  ],
  "query_params": {
    "hello": "woeoio"
  }
}
17:26:16.500 Debug Info:   {
  "Request": {
    "IsAsync": false,
    "Method": "POST",
    "Url": "https://vb6.pro/echo/?hello=woeoio",
    "Body": "{\"consignee\":{\"address\":{\"city\":\"Berlin\",\"countryCode\":\"DE\",\"postCode\":\"10115\"}},\"lines\":[{\"content\":\"furniture\",\"unitHeight\":120,\"unitLength\":120,\"unitPackageType\":\"EP\",\"unitQuantity\":1,\"unitWeight\":200,\"unitWidth\":80}],\"product\":\"STANDARD\",\"pickupOn\":\"2026-02-08\"}",
    "Headers": {
      "X-API-KEY": "YOUR API KEY HERE",
      "Content-Type": "application/json"
    },
    "TimeOut": 5,
    "ChartSet": "utf-8"
  },
  "Response": {
    "Status": 200,
    "StatusText": "OK",
    "Headers": {
      "Cache-Control": "no-store, no-cache, must-revalidate",
      "Connection": "keep-alive",
      "Date": "Sun, 08 Feb 2026 09",
      "Pragma": "no-cache",
      "Transfer-Encoding": "chunked",
      "Content-Type": "application/json; charset=utf-8",
      "Expires": "Thu, 19 Nov 1981 08",
      "Server": "nginx",
      "Set-Cookie": "PHPSESSID=ngio70673jbqcsstl2f86bck89; path=/",
      "Access-Control-Allow-Origin": "*",
      "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, PATCH, OPTIONS, HEAD",
      "Access-Control-Allow-Headers": "Content-Type, Authorization, X-Requested-With",
      "Strict-Transport-Security": "max-age=31536000",
      "": ""
    },
    "Content": "{\n    \"server\": {\n        \"time\": \"2026-02-08 17:26:14 CST\",\n        \"timestamp\": 1770542774,\n        \"timezone\": \"PRC\",\n        \"software\": \"nginx/1.22.1\",\n        \"name\": \"www.vb6.pro\",\n        \"address\": \"172.23.151.44\",\n        \"port\": \"443\",\n        \"protocol\": \"HTTP/1.1\",\n        \"document_root\": \"/www/wwwroot/vb6.pro\",\n        \"php_version\": \"8.2.28\",\n        \"sapi\": \"fpm-fcgi\",\n        \"os\": \"Linux\"\n    },\n    \"request\": {\n        \"method\": \"POST\",\n        \"uri\": \"/echo/?hello=woeoio\",\n        \"url\": \"https://\",\n        \"query_string\": \"hello=woeoio\",\n        \"path_info\": \"\",\n        \"script_name\": \"/echo/index.php\",\n        \"filename\": \"/www/wwwroot/vb6.pro/echo/index.php\",\n        \"time\": \"2026-02-08 17:26:14 CST\",\n        \"timestamp\": 1770542774,\n        \"time_float\": 1770542774.17496,\n        \"http_version\": \"HTTP/1.1\",\n        \"is_ajax\": false,\n        \"is_secure\": true,\n        \"full_url\": \"https://vb6.pro/echo/?hello=woeoio\",\n        \"base_url\": \"https://vb6.pro/echo\"\n    },\n    \"client\": {\n        \"ip\": \"61.146.189.212\",\n        \"port\": \"21702\",\n        \"user_agent\": \"Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)\",\n        \"referer\": \"\",\n        \"language\": \"\",\n        \"encoding\": \"\",\n        \"charset\": \"\",\n        \"connection\": \"Keep-Alive\",\n        \"host\": \"vb6.pro\"\n    },\n    \"headers\": {\n        \"Host\": \"vb6.pro\",\n        \"Content-Length\": \"266\",\n        \"Content-Lenght\": \"266\",\n        \"X-Api-Key\": \"YOUR API KEY HERE\",\n        \"User-Agent\": \"Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)\",\n        \"Accept\": \"*/*\",\n        \"Content-Type\": \"application/json\",\n        \"Connection\": \"Keep-Alive\"\n    },\n    \"get\": {\n        \"hello\": \"woeoio\"\n    },\n    \"post\": {\n        \"consignee\": {\n            \"address\": {\n                \"city\": \"Berlin\",\n                \"countryCode\": \"DE\",\n                \"postCode\": \"10115\"\n            }\n        },\n        \"lines\": [\n            {\n                \"content\": \"furniture\",\n                \"unitHeight\": 120,\n                \"unitLength\": 120,\n                \"unitPackageType\": \"EP\",\n                \"unitQuantity\": 1,\n                \"unitWeight\": 200,\n                \"unitWidth\": 80\n            }\n        ],\n        \"product\": \"STANDARD\",\n        \"pickupOn\": \"2026-02-08\"\n    },\n    \"files\": [],\n    \"cookies\": [],\n    \"session\": [],\n    \"body\": \"{\\\"consignee\\\":{\\\"address\\\":{\\\"city\\\":\\\"Berlin\\\",\\\"countryCode\\\":\\\"DE\\\",\\\"postCode\\\":\\\"10115\\\"}},\\\"lines\\\":[{\\\"content\\\":\\\"furniture\\\",\\\"unitHeight\\\":120,\\\"unitLength\\\":120,\\\"unitPackageType\\\":\\\"EP\\\",\\\"unitQuantity\\\":1,\\\"unitWeight\\\":200,\\\"unitWidth\\\":80}],\\\"product\\\":\\\"STANDARD\\\",\\\"pickupOn\\\":\\\"2026-02-08\\\"}\",\n    \"parsed_body\": {\n        \"json\": {\n            \"consignee\": {\n                \"address\": {\n                    \"city\": \"Berlin\",\n                    \"countryCode\": \"DE\",\n                    \"postCode\": \"10115\"\n                }\n            },\n            \"lines\": [\n                {\n                    \"content\": \"furniture\",\n                    \"unitHeight\": 120,\n                    \"unitLength\": 120,\n                    \"unitPackageType\": \"EP\",\n                    \"unitQuantity\": 1,\n                    \"unitWeight\": 200,\n                    \"unitWidth\": 80\n                }\n            ],\n            \"product\": \"STANDARD\",\n            \"pickupOn\": \"2026-02-08\"\n        }\n    },\n    \"content\": {\n        \"type\": \"application/json\",\n        \"length\": \"266\",\n        \"md5\": \"fbf19d79627687cd67ddd6eaa2971f43\",\n        \"encoding\": \"\"\n    },\n    \"auth\": {\n        \"type\": \"\",\n        \"user\": \"\"\n    },\n    \"ssl\": {\n        \"protocol\": \"Unknown\",\n        \"cipher\": \"Unknown\",\n        \"version\": \"Unknown\",\n        \"verify_client\": \"NONE\"\n    },\n    \"phpinfo\": [],\n    \"query_params\": {\n        \"hello\": \"woeoio\"\n    }\n}"
  }
}
17:26:16.525 (time taken: 1.9914793s)
```
