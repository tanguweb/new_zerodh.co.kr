<!-- #include virtual = "/Include/Config.asp" -->
<%
	'########################################################################################
	'#	File		: /manager/logout.asp
	'#  Create		: domoyosi / 2015.06.30
	'#	Info		: ¾îµå¹Î ·Î±×¾Æ¿ôÃ³¸®
	'#	Update		: 
	'#	Update Memo	: 
	'########################################################################################

	Response.Cookies("member")					= ""
	Response.Cookies("member")("m_UserID")		= ""
	Response.Cookies("member")("m_UserNM")		= ""
	'Response.Cookies("member")("m_PrimaryID")	= ""
	Response.Cookies("member")("m_Admin")		= ""
	Response.Cookies("member").Path = "/"
	Response.Cookies("member").domain = CookieDomain

	response.write "<script>document.location.href='/manager/login.asp';</script>"
	response.end 
%>