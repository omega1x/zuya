'Simple demonstration of geo-coding in ZULUGIS using Yandex geo-coder API
'2020 © Yuri Possokhov [PosokhovIu@sibgenco.ru]
'Siberian Generating Company, LLC

Const YANDEX_API_KEY = "?" 'change this Yandex-key to your own

Sub MACRO_YANDEX_GEOCODING
  ' 1. Get Yandex representation of some address and its coordinates:   
   Some_Address = "Новосибирск,  Федора Ивачева ул. , 4" 
  ' - Ask Yandex:   
   Yandex_response = ya_code(URLEncode(Some_Address), YANDEX_API_KEY)
   
  ' - Process Yandex Response:   
   Yandex_Address = get_address(Yandex_response)
   Yandex_Lat = get_latitude(Yandex_response)
   Yandex_Lon = get_longitude(Yandex_response)   
   MsgBox(Yandex_Address + "|" + Yandex_Lat + "|" + Yandex_Lon)
      
  ' 2. Get Yandex address using coordinates only:
   Yandex_response = ya_coord(55.760241, 37.611347, YANDEX_API_KEY)
   
  ' - Process Yandex Response:   
   Yandex_Address = get_address(Yandex_response)
   Yandex_Lat = get_latitude(Yandex_response)
   Yandex_Lon = get_longitude(Yandex_response)   
   
   MsgBox(Yandex_Address + "|" + Yandex_Lat + "|" + Yandex_Lon)
End Sub

Function HTTPReq(query)
  'Send http-request to server
   Set Req = CreateObject("WinHttp.WinHttpRequest.5.1")
   Req.Option(4) = 13056 '
   Req.Open "GET", query, False
   Req.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
   Req.send ("")
   HTTPReq = Req.responseText
End Function

Function ya_code(address, api_key) 
  'Get Yandex representation of some address and its coordinates
   Const yandex_service = "https://geocode-maps.yandex.ru/1.x"
   query = yandex_service + "/?" + "apikey=" + api_key + "&geocode=" + address
   Set XML = CreateObject("Microsoft.XMLDOM")
   XML.LoadXML(HTTPReq(query))
   ya_code = XML.getElementsByTagName("formatted")(0).ChildNodes(0).nodeValue & "|" & XML.getElementsByTagName("pos")(0).ChildNodes(0).nodeValue 
End Function

Function ya_coord(lat, lon, api_key)
  'Get Yandex address for given coordinates
   Const yandex_service = "https://geocode-maps.yandex.ru/1.x"
   query = yandex_service + "/?" + "apikey=" + api_key + "&geocode=" & lon  & "," & lat & "&results=1"
   Set XML = CreateObject("Microsoft.XMLDOM")
   XML.LoadXML(HTTPReq(query))
   ya_coord = XML.getElementsByTagName("text")(0).ChildNodes(0).nodeValue + "|" + XML.getElementsByTagName("pos")(0).ChildNodes(0).nodeValue
End Function

Function get_address(text)
    get_address = Split(text, "|")(0)
End Function

Function get_latitude(text)
    get_latitude = Split(Split(text, "|")(1))(1)
End Function

Function get_longitude(text)
    get_longitude = Split(Split(text, "|")(1))(0)
End Function

Function URLEncode(ByVal txt)
'Encode Russian symbols for URL
 For i = 1 To Len(txt)
	l = Mid(txt, i, 1)
		
        If AscW(l) > 4095 Then 
            t = "%" & Hex(AscW(l) \ 64 \ 64 + 224) & "%" & Hex(AscW(l) \ 64) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
        ElseIf AscW(l) > 127 Then          
            t = "%" & Hex(AscW(l) \ 64 + 192) & "%" & Hex(8 * 16 + AscW(l) Mod 64)
        ElseIf AscW(l) = 32 Then
            t = "%20"
        Else 
            t = l
        End If
        URLEncode = URLEncode & t
    Next
End Function
