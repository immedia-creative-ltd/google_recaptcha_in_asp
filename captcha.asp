<script src="https://www.google.com/recaptcha/api.js" async defer></script>

<%
'Response.write Cat_ID
'Response.write "<br />sv "
'Response.write Request.ServerVariables("REQUEST_METHOD")
'Response.write "<br />"

Dim name, email, confirm_email, telephone, message, outcome
name = request.form("name")
email = request.form("email")
confirm_email = request.form("confirm_email")
telephone = request.form("telephone")
message = request.form("message")
formsubmitted = request.form("formsubmitted")
outcome = request.form("outcome")

'Response.write "Name is " & name
'Response.write "<br />email is " & email
'Response.write "<br />confirm_email is " & confirm_email
'Response.write "<br />Telephone is " & telephone
'Response.write "<br />Message is " & message
'Response.write "<br />formsubmitted is " & formsubmitted
'Response.write "<br />Outcome is " & outcome

	'If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
	If formsubmitted = "yes" Then
		
		if name = "" then
			Response.write "<script language='javascript'>"
			Response.write "alert('Please provide your name')"
			Response.write "</script>"
		end if

		if email = "" then
			Response.write "<script language='javascript'>"
			Response.write "alert('Please provide your email address')"
			Response.write "</script>"
		end if

		if email <> confirm_email then
			Response.write "<script language='javascript'>"
			Response.write "alert('Your email addresses do not match')"
			Response.write "</script>"
		end if

		if name <> "" and email <> "" and confirm_email <> "" then
		%>
			<script type="text/javascript">
			function formAutoSubmit () {
			var frm = document.getElementById("forwardForm");
			frm.submit();
			}
			window.onload = formAutoSubmit;
			</script>
		<%
		end if

		Dim recaptcha_secret, sendstring, objXML
		' Secret key
		recaptcha_secret = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"

		sendstring = "https://www.google.com/recaptcha/api/siteverify?secret=" & recaptcha_secret & "&response=" & Request.form("g-recaptcha-response")

		Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")
		objXML.Open "GET", sendstring, False

		objXML.Send

		'Response.write "<h3>Response: " & objXML.responseText & "</h3>"

		Dim ResponseString
		ResponseString = split(objXML.responseText, vblf)

		Set objXML = Nothing

		If InStr(ResponseString(1), "true") = 0 Then
			'They answered incorrectly
			'Response.write "<h3>success = false</h3>"
		else
			'They answered correctly
			'Response.write InStr(ResponseString(1), "true")
			'Response.write "<h3>success = true</h3>"
			%>
			<form method="post" id="forwardForm" action="emailme.asp">
			<!--<form method="post" id="forwardForm" action="emailme.asp">-->
				<input type="hidden" name="name" value="<%=name%>">
				<input type="hidden" name="email" value="<%=email%>">
				<input type="hidden" name="telephone" value="<%=telephone%>">
				<input type="hidden" name="message" value="<%=message%>">
				<input type="hidden" name="confirm_email" value="<%=confirm_email%>">
				<input type="hidden" name="outcome" value="form submitted">
			</form>
			<%
		end if
  
    End If
%>
<h1>How to Contact Solenvis</h1>
<p>Solenvis Limited, Unit 3, Old Farm Shop, Howe Lane Farm Estate, Howe Lane, White Waltham, Maidenhead, Berkshire, SL6 3JP</p>
<p>Tel: 01189 342607</p>
<p>Email: <a href="&#109;&#97;&#105;&#108;&#116;&#111;&#58;&#115;&#97;&#108;&#101;&#115;&#64;&#115;&#111;&#108;&#101;&#110;&#118;&#105;&#115;&#102;&#108;&#111;&#119;&#109;&#101;&#116;&#101;&#114;&#115;&#46;&#99;&#111;&#109;">sales@solenvisflowmeters.com</a></p>
<p><em>Alternatively, you may wish to complete and send the following form with your details and request, and a member of our sales team will call you back.&nbsp; Please note when filling in the form that you need to fill in all fields and omit spaces in the phone number. When filling the security box include upper and lower case characters and&nbsp; punctuation as shown. If successful a thank you message will appear. </em></p>
<div style="float:left;">

        <form method="POST" action="" id="userForm">
			<div class="FormTable">
			<div class="FormTableBody">
			<div class="FormTableRow">
			<div class="FormTableCell">Name*&nbsp;</div>
			<div class="FormTableCellr"><input type="text" name="name" size="30" value="<%=name%>" class="search" style="width: 170px;" /></div>
			</div>
			<div class="FormTableRow">
			<div class="FormTableCell">Email *&nbsp;</div>
			<div class="FormTableCellr"><input type="text" name="email" size="30" value="<%=email%>" class="search" style="width: 170px;" /></div>
			</div>
			<div class="FormTableRow">
			<div class="FormTableCell">Confirm Email *&nbsp;</div>
			<div class="FormTableCellr"><input type="text" name="confirm_email" size="30" value="<%=confirm_email%>" class="search" style="width: 170px;" /></div>
			</div>
			<div class="FormTableRow">
			<div class="FormTableCell">Telephone</div>
			<div class="FormTableCellr"><input type="text" maxlength="20" name="telephone" size="30" value="<%=telephone%>" class="search" style="width: 170px;" /></div>
			</div>
			<div class="FormTableRow">
			<div class="FormTableCell">Message</div>
			<div class="FormTableCellr"><textarea cols="22" name="message" rows="2" class="textarea"><%=message%></textarea></div>
			</div>
			<div class="FormTableRow">
			</div>
			</div>
			</div>
			<input type="hidden" value="yes" name="formsubmitted" />
			<input type="hidden" value="contact-confirmation" name="PageRedirect" />
            <button class="g-recaptcha btn btn-primary" data-sitekey="yyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyyy" data-callback="submitForm">Submit</button>
        </form>

		<script>
		function submitForm() {
			document.getElementById("userForm").submit();
		}
		</script>
<p>&nbsp;</p>	
</div>
<p><iframe allowfullscreen="" style="border:0" src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d2484.9798927761512!2d-0.7632560842311585!3d51.47688337963003!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x0%3A0xc06249062d07adae!2sSolenvis+Ltd!5e0!3m2!1sen!2sus!4v1461773910008" width="425" height="350" frameborder="0"></iframe></p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p><span style="border-radius: 2px; text-indent: 20px; width: auto; padding: 0px 4px 0px 0px; text-align: center; font: bold 11px/20px &quot;Helvetica Neue&quot;,Helvetica,sans-serif; color: #ffffff; background: #bd081c url(&quot;data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIGhlaWdodD0iMzBweCIgd2lkdGg9IjMwcHgiIHZpZXdCb3g9Ii0xIC0xIDMxIDMxIj48Zz48cGF0aCBkPSJNMjkuNDQ5LDE0LjY2MiBDMjkuNDQ5LDIyLjcyMiAyMi44NjgsMjkuMjU2IDE0Ljc1LDI5LjI1NiBDNi42MzIsMjkuMjU2IDAuMDUxLDIyLjcyMiAwLjA1MSwxNC42NjIgQzAuMDUxLDYuNjAxIDYuNjMyLDAuMDY3IDE0Ljc1LDAuMDY3IEMyMi44NjgsMC4wNjcgMjkuNDQ5LDYuNjAxIDI5LjQ0OSwxNC42NjIiIGZpbGw9IiNmZmYiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIxIj48L3BhdGg+PHBhdGggZD0iTTE0LjczMywxLjY4NiBDNy41MTYsMS42ODYgMS42NjUsNy40OTUgMS42NjUsMTQuNjYyIEMxLjY2NSwyMC4xNTkgNS4xMDksMjQuODU0IDkuOTcsMjYuNzQ0IEM5Ljg1NiwyNS43MTggOS43NTMsMjQuMTQzIDEwLjAxNiwyMy4wMjIgQzEwLjI1MywyMi4wMSAxMS41NDgsMTYuNTcyIDExLjU0OCwxNi41NzIgQzExLjU0OCwxNi41NzIgMTEuMTU3LDE1Ljc5NSAxMS4xNTcsMTQuNjQ2IEMxMS4xNTcsMTIuODQyIDEyLjIxMSwxMS40OTUgMTMuNTIyLDExLjQ5NSBDMTQuNjM3LDExLjQ5NSAxNS4xNzUsMTIuMzI2IDE1LjE3NSwxMy4zMjMgQzE1LjE3NSwxNC40MzYgMTQuNDYyLDE2LjEgMTQuMDkzLDE3LjY0MyBDMTMuNzg1LDE4LjkzNSAxNC43NDUsMTkuOTg4IDE2LjAyOCwxOS45ODggQzE4LjM1MSwxOS45ODggMjAuMTM2LDE3LjU1NiAyMC4xMzYsMTQuMDQ2IEMyMC4xMzYsMTAuOTM5IDE3Ljg4OCw4Ljc2NyAxNC42NzgsOC43NjcgQzEwLjk1OSw4Ljc2NyA4Ljc3NywxMS41MzYgOC43NzcsMTQuMzk4IEM4Ljc3NywxNS41MTMgOS4yMSwxNi43MDkgOS43NDksMTcuMzU5IEM5Ljg1NiwxNy40ODggOS44NzIsMTcuNiA5Ljg0LDE3LjczMSBDOS43NDEsMTguMTQxIDkuNTIsMTkuMDIzIDkuNDc3LDE5LjIwMyBDOS40MiwxOS40NCA5LjI4OCwxOS40OTEgOS4wNCwxOS4zNzYgQzcuNDA4LDE4LjYyMiA2LjM4NywxNi4yNTIgNi4zODcsMTQuMzQ5IEM2LjM4NywxMC4yNTYgOS4zODMsNi40OTcgMTUuMDIyLDYuNDk3IEMxOS41NTUsNi40OTcgMjMuMDc4LDkuNzA1IDIzLjA3OCwxMy45OTEgQzIzLjA3OCwxOC40NjMgMjAuMjM5LDIyLjA2MiAxNi4yOTcsMjIuMDYyIEMxNC45NzMsMjIuMDYyIDEzLjcyOCwyMS4zNzkgMTMuMzAyLDIwLjU3MiBDMTMuMzAyLDIwLjU3MiAxMi42NDcsMjMuMDUgMTIuNDg4LDIzLjY1NyBDMTIuMTkzLDI0Ljc4NCAxMS4zOTYsMjYuMTk2IDEwLjg2MywyNy4wNTggQzEyLjA4NiwyNy40MzQgMTMuMzg2LDI3LjYzNyAxNC43MzMsMjcuNjM3IEMyMS45NSwyNy42MzcgMjcuODAxLDIxLjgyOCAyNy44MDEsMTQuNjYyIEMyNy44MDEsNy40OTUgMjEuOTUsMS42ODYgMTQuNzMzLDEuNjg2IiBmaWxsPSIjYmQwODFjIj48L3BhdGg+PC9nPjwvc3ZnPg==&quot;) no-repeat scroll 3px 50% / 14px 14px; position: absolute; opacity: 1; z-index: 8675309; display: none; cursor: pointer; border: medium none;">Save</span></p>
