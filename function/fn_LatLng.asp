<%
	'// 주소(address)를 파라미터로 위도, 경도 값을 얻어 옵니다.  
	'// Lat(위도), Lng(경도)
	'// 리턴값 = 경도, 위도
	'// split(리턴값, ",") 으로 분리 시켜 각각 저장 합니다.
	Function setLatLng(address)

		Dim NaverKey, keyValue
		Dim xml, dom, node1, node2, msg

		On Error Resume Next
			'// Naver Map API Key 입니다. (http://peters.webizsolution.co.kr:8090/ 하위 디렉토리 모두 적용 됩니다.)
			NaverKey = "c51de9390acda6e706f7af1c15e18937"

			'## 좌표를 가져올 주소를 네이년에게 던집니다.
			'## 1. Naver 에 주소를 파라미터로 전달하는 URL (http://maps.naver.com/api/geocode.php)
			'## 2. coord : tm128(X,Y 6자리 좌표계), latlng(위도/경도)
			keyValue = "http://maps.naver.com/api/geocode.php?key=" & NaverKey & "&encoding=utf-8&coord=latlng"  & "&query=" & Server.UrlEncode(address)	'address : 주소 파라미터

			'## keyValue 
			Set xml = Server.CreateObject("Msxml2.ServerXMLHTTP")
			xml.open "GET", keyValue, False
			xml.send
			
			'##
			Set dom = Server.CreateObject("Microsoft.XMLDOM")
			dom.async = False
			dom.load xml.responseBody

			'## xml 의 x 좌표를 가져 옵니다.
			Set node1 = dom.documentElement.SelectSingleNode("/geocode/item/point/x")
			msg = node1.text
			
			'## xml 의 y 좌표를 가져와서 msg 에 담습니다.
			Set node2 = dom.documentElement.SelectSingleNode("/geocode/item/point/y")
			msg = msg & "," & node2.text

			'## 에러처리
			If Err.Number <> 0 Then
				response.write "<font size='2pt'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;정보가 올바르지 않은 경우 지도가 안 보일 수도 있습니다.<font>"
				'response.write "<br>Err.Number 오류 번호 " & Err.Number & " : " & Err.Description
				response.End
			End If
				setLatLng = msg

				

	End Function
%>