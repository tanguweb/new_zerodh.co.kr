<%
	function fn_XMLfilepath(boardcd, seq)
		dim rtn_path
		select case boardcd
			case 1010
				if seq = 1 then 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\image\thumb\"
				elseif seq = 2 then 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\image\big\"
				elseif seq = 3 then 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\image\large\"
				elseif seq = 4 then 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\image\"
				else 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\cadalog\image\"
				end If
			Case 1011
				If seq = 1 Then
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\coordi\flash_image\listimg\"
				ElseIf seq = 2 Then
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\coordi\flash_image\bigimg\"
				Else 
					rtn_path = "D:\Project\[inthef]www.inthef.co.kr\joinus\_swf\style\coordi\flash_image\"
				End If
			case else
				rtn_path = ""
		end select
		fn_XMLfilepath = rtn_path
	end function 

	function fn_XMLviewpath(boardcd, seq)
		dim rtn_path
		select case boardcd
			case 1010
				if seq = 1 then 
					rtn_path = "/joinus/_swf/style/cadalog/image/thumb/"
				elseif seq = 2 then 
					rtn_path = "/joinus/_swf/style/cadalog/image/big/"
				elseif seq = 3 then 
					rtn_path = "/joinus/_swf/style/cadalog/image/large/"
				elseif seq = 4 then 
					rtn_path = "/joinus/_swf/style/cadalog/image/"
				else 
					rtn_path = "/joinus/_swf/style/cadalog/image/"
				end If
			Case 1011
				If seq = 1 Then
					rtn_path = "/joinus/_swf/style/coordi/flash_image/listimg/"
				ElseIf seq = 2 Then
					rtn_path = "/joinus/_swf/style/coordi/flash_image/bigimg/"
				Else 
					rtn_path = "/joinus/_swf/style/coordi/flash_image/"
				End If
			case else
				rtn_path = ""
		end select
		fn_XMLviewpath = rtn_path
	end function 
%>