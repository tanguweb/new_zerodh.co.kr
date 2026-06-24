function GetHttpRequestObject() 
{
	var httpReq;
	try 
	{		
		if (window.ActiveXObject)		//Non-IE browsers
		{
			httpReq = new ActiveXObject("Microsoft.XMLHTTP");
		}
		else if (window.XMLHttpRequest) //IE
		{ 
			httpReq = new XMLHttpRequest();
		}
		return httpReq;  
	}
	catch(e) 
	{
		//alert("GetHttpRequestObject: " + e.name + ":" + e.message);
		return null;
	}
}//GetHttpRequestObject()

function HttpRequest()
{
	//기본정보
	this.version = "1.0.0612151434";
	this.name = "HttpRequest";
	this.resultText = "default";

	//셋팅해야할 값
	this.url	= "";
	this.method = "GET"; //default value

	this.response = function()
	{
		try
		{
			this.method = this.method.toUpperCase();
			if(this.url == "")
			{
				this.resultText = "error:url";			
				return "error:url";
			}
			//define cursor type
			document.body.style.cursor = "wait";			
			//create object
			var oXMLHTTP = GetHttpRequestObject();

			if (oXMLHTTP == null)
			{
				document.body.style.cursor = "auto";
				return "";
			}

			//Send
			oXMLHTTP.open(this.method, this.url, false);
			if(this.method == "POST")
			{
				oXMLHTTP.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
			}
			oXMLHTTP.send();			
			//receive data
			this.resultText = oXMLHTTP.responseText;
			//define cursor type
			document.body.style.cursor = "auto";
			//return result
			return this.resultText;
		}
		catch(e)
		{
			//alert("HttpRequest: " + e.name + ":" + e.message);
			document.body.style.cursor = "auto";
			return "error:object"
		}
	}
}//HttpRequest()


function HttpRequest2()
{
	//기본정보
	this.version = "1.1.0701241152";
	this.name = "HttpRequest";
	this.resultText = "default";

	//셋팅해야할 값
	this.method = "GET"; //default value
	this.url	= "";
	this.params = "";
	this.async = false; //비동기호출여부
	this.responsetype = "TEXT";

	this.response = function()
	{
		try
		{
			this.method = this.method.toUpperCase();
			if(this.url == "")
			{
				this.resultText = "error:url";			
				return "error:url";
			}
			//define cursor type
			document.body.style.cursor = "wait";			
			//create object
			var oXMLHTTP = GetHttpRequestObject();

			if (oXMLHTTP == null)
			{
				document.body.style.cursor = "auto";
				alert("고객님의 브라우즈가 Ajax를 지원하지 않습니다.");
				return "";
			}
			
			if(this.method == "GET")
			{
				//Send
				oXMLHTTP.open(this.method, this.url + "?" + this.params , this.async);
				oXMLHTTP.send();
			}
			else if(this.method == "POST")
			{
				oXMLHTTP.open(this.method, this.url, this.async);
				oXMLHTTP.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
				oXMLHTTP.send(this.params);
			}
			else
			{
				document.body.style.cursor = "auto";
				alert("전송메소드가 올바르지 않습니다.[ GET || POST ]");
				return "";
			}
			//receive data
			this.resultText = oXMLHTTP.responseText;
			//define cursor type
			document.body.style.cursor = "auto";
			//return result
			return this.resultText;
		}
		catch(e)
		{
			//alert("HttpRequest: " + e.name + ":" + e.message);
			document.body.style.cursor = "auto";
			return "error:object"
		}
	}
}//HttpRequest()