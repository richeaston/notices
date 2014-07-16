<div class="bd">
<form action="<%= ew_CurrentPage %>">
<input type="hidden" name="export" id="export" value="email">
<table border="0" cellspacing="0" cellpadding="4">
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormSender") %></span></td>
		<td><span class="aspmaker"><input type="text" name="sender" id="sender" size="30"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormRecipient") %></span></td>
		<td><span class="aspmaker"><input type="text" name="recipient" id="recipient" size="30"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormCc") %></span></td>
		<td><span class="aspmaker"><input type="text" name="cc" id="cc" size="30"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormBcc") %></span></td>
		<td><span class="aspmaker"><input type="text" name="bcc" id="bcc" size="30"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormSubject") %></span></td>
		<td><span class="aspmaker"><input type="text" name="subject" id="subject" size="50"></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormMessage") %></span></td>
		<td><span class="aspmaker"><textarea cols="50" rows="8" name="message" id="message"></textarea></span></td>
	</tr>
	<tr>
		<td><span class="aspmaker"><%= Language.Phrase("EmailFormContentType") %></span></td>
		<td><span class="aspmaker">
		<label><input type="radio" name="contenttype" id="contenttype" value="html" checked="checked"><%= Language.Phrase("EmailFormContentTypeHtml") %></label>
		<label><input type="radio" name="contenttype" id="contenttype" value="url"><%= Language.Phrase("EmailFormContentTypeUrl") %></label>
		</span></td>
	</tr>
</table>
</form>
</div>
