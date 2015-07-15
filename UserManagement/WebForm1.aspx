<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="UserManagement.WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
  <head>
    <title>WebForm1</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name=vs_defaultClientScript content="JavaScript">
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">

  </head>
  
  <body MS_POSITIONING="GridLayout">



    <form id="form1" method="post" runat="server">
<input type='text'  id="adriano" onkeyup="fnAlphaNumericOnly(this)"><input type="button" onclick="showKeyCode()">



<script language = 'javascript'>  

function fnAlphaNumericOnly()
{
		var txtbox = form1.adriano
		var ToTest = txtbox.value
		//var txtbox = document.getElementByID("adriano")
		var re = /[^0-9A-Za-z]/
		if (re.test(ToTest)) {
			txtbox.value = txtbox.value.substring(0, txtbox.value.length-1)
		}
}

/*
function showKeyCode(txtbox)
{
			var AsciiCode
			var Text = txtbox.value
			for (i=0;i<Text.length;i++)  {
				AsciiCode = Text.charCodeAt(i)
			         if ((charCode > 31 && charCode < 48 ) || (charCode > 57  && charCode < 65 ) || (charCode > 90  && charCode < 97) || (charCode > 122))	  
			
				//Char = Text.charAt(i)
				//if (ValidChars.indexOf(Char) == -1)  {
				//	IsValid = false
				//}	
			}
         if ((charCode > 31 && charCode < 48 ) || (charCode > 57  && charCode < 65 ) || (charCode > 90  && charCode < 97) || (charCode > 122))
            return false;

		var character = txtbox.value
		var asciicode = character.charCodeAt(0)
		
		
		alert(msg);		
}
*/
</script>
    </form>
    
    

  </body>
</html>
