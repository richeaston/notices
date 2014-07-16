// Check US Date format (mm/dd/yyyy)
function ew_CheckUSDate(object_value) {
	return ew_CheckDateEx(object_value, "us", EW_DATE_SEPARATOR);
}

// Check US Date format (mm/dd/yy)
function ew_CheckShortUSDate(object_value) {
	return ew_CheckDateEx(object_value, "usshort", EW_DATE_SEPARATOR);
}

// Check Date format (yyyy/mm/dd)
function ew_CheckDate(object_value) {
	return ew_CheckDateEx(object_value, "std", EW_DATE_SEPARATOR);
}

// Check Date format (yy/mm/dd)
function ew_CheckShortDate(object_value) {
	return ew_CheckDateEx(object_value, "stdshort", EW_DATE_SEPARATOR);
}

// Check Euro Date format (dd/mm/yyyy)
function ew_CheckEuroDate(object_value) {
	return ew_CheckDateEx(object_value, "euro", EW_DATE_SEPARATOR);
}

// Check Euro Date format (dd/mm/yy)
function ew_CheckShortEuroDate(object_value) {
	return ew_CheckDateEx(object_value, "euroshort", EW_DATE_SEPARATOR);
}

// Check date format
// format: std/stdshort/us/usshort/euro/euroshort

function ew_CheckDateEx(value, format, sep) {
	if (value == null || value.length == "")
		return true;
	while (value.indexOf("  ") > -1)
		value = value.replace(/  /g, " ");
	value = value.replace(/^\s*|\s*$/g, "");
	var arDT = value.split(" ");
	if (arDT.length > 0) {
		var re, sYear, sMonth, sDay;
		re = /^([0-9]{4})-([0][1-9]|[1][0-2])-([0][1-9]|[1|2][0-9]|[3][0|1])$/;
		if (ar = re.exec(arDT[0])) {
			sYear = ar[1];
			sMonth = ar[2];
			sDay = ar[3];
		} else {
			var wrksep = "\\" + sep;
			switch (format) {
				case "std":
					re = new RegExp("^([0-9]{4})" + wrksep + "([0]?[1-9]|[1][0-2])" + wrksep + "([0]?[1-9]|[1|2][0-9]|[3][0|1])$");
					break;
				case "stdshort":
					re = new RegExp("^([0-9]{2})" + wrksep + "([0]?[1-9]|[1][0-2])" + wrksep + "([0]?[1-9]|[1|2][0-9]|[3][0|1])$");
					break;
				case "us":
					re = new RegExp("^([0]?[1-9]|[1][0-2])" + wrksep + "([0]?[1-9]|[1|2][0-9]|[3][0|1])" + wrksep + "([0-9]{4})$");
					break;
				case "usshort":
					re = new RegExp("^([0]?[1-9]|[1][0-2])" + wrksep + "([0]?[1-9]|[1|2][0-9]|[3][0|1])" + wrksep + "([0-9]{2})$");
					break;
				case "euro":
					re = new RegExp("^([0]?[1-9]|[1|2][0-9]|[3][0|1])" + wrksep + "([0]?[1-9]|[1][0-2])" + wrksep + "([0-9]{4})$");
					break;
				case "euroshort":
					re = new RegExp("^([0]?[1-9]|[1|2][0-9]|[3][0|1])" + wrksep + "([0]?[1-9]|[1][0-2])" + wrksep + "([0-9]{2})$");
					break;
			}
			if (!re.test(arDT[0]))
				return false;
		}
		var arD = arDT[0].split(sep);
		switch (format) {
			case "std":
			case "stdshort":
				sYear = ew_UnformatYear(arD[0]);
				sMonth = arD[1];
				sDay = arD[2];
				break;
			case "us":
			case "usshort":
				sYear = ew_UnformatYear(arD[2]);
				sMonth = arD[0];
				sDay = arD[1];
				break;
			case "euro":
			case "euroshort":
				sYear = ew_UnformatYear(arD[2]);
				sMonth = arD[1];
				sDay = arD[0];
				break;
		}
		if (!ew_CheckDay(sYear, sMonth, sDay))
			return false;
	}
	if (arDT.length > 1 && !ew_CheckTime(arDT[1]))
		return false;
	return true;
}

// Unformat 2 digit year to 4 digit year
function ew_UnformatYear(yr) {
	if (yr.length == 2) {
		if (yr > EW_UNFORMAT_YEAR)
			return "19" + yr;
		else
			return "20" + yr;
	} else {
		return yr;
	}
}

// Check day
function ew_CheckDay(checkYear, checkMonth, checkDay) {
	maxDay = 31;
	if (checkMonth == 4 || checkMonth == 6 ||	checkMonth == 9 || checkMonth == 11) {
		maxDay = 30;
	} else if (checkMonth == 2)	{
		if (checkYear % 4 > 0)
			maxDay =28;
		else if (checkYear % 100 == 0 && checkYear % 400 > 0)
			maxDay = 28;
		else
			maxDay = 29;
	}
	return ew_CheckRange(checkDay, 1, maxDay);
}

// Check integer
function ew_CheckInteger(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	var decimal_format = ".";
	var check_char;
	check_char = object_value.indexOf(decimal_format);
	if (check_char < 1)
		return ew_CheckNumber(object_value);
	else
		return false;
}

// Check number range
function ew_NumberRange(object_value, min_value, max_value) {
	if (min_value != null) {
		if (object_value < min_value)
			return false;
	}
	if (max_value != null) {
		if (object_value > max_value)
			return false;
	}
	return true;
}

// Check number
function ew_CheckNumber(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	var start_format = " .+-0123456789";
	var number_format = " .0123456789";
	var check_char;
	var decimal = false;
	var trailing_blank = false;
	var digits = false;
	check_char = start_format.indexOf(object_value.charAt(0));
	if (check_char == 1)
		decimal = true;
	else if (check_char < 1)
		return false;
	for (var i = 1; i < object_value.length; i++)	{
		check_char = number_format.indexOf(object_value.charAt(i))
		if (check_char < 0) {
			return false;
		} else if (check_char == 1)	{
			if (decimal)
				return false;
			else
				decimal = true;
		} else if (check_char == 0) {
			if (decimal || digits)	
			trailing_blank = true;
		}	else if (trailing_blank) { 
			return false;
		} else {
			digits = true;
		}
	}	
	return true;
}

// Check range
function ew_CheckRange(object_value, min_value, max_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	if (!ew_CheckNumber(object_value))
		return false;
	else
		return (ew_NumberRange((eval(object_value)), min_value, max_value));
	return true;
}

// Check time
function ew_CheckTime(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
    var time_value;
    if (object_value.indexOf('.') > 0) {
        time_value = object_value.substring(0, object_value.indexOf('.'));
    } else if (object_value.indexOf(' ') > 0) {
        time_value = object_value.substring(0, object_value.indexOf(' '));
    } else {
        time_value = object_value;
    }
	isplit = time_value.indexOf(':');
	if (isplit == -1 || isplit == time_value.length)
		return false;
	sHour = time_value.substring(0, isplit);
	iminute = time_value.indexOf(':', isplit + 1);
	if (iminute == -1 || iminute == time_value.length)
		sMin = time_value.substring((sHour.length + 1));
	else
		sMin = time_value.substring((sHour.length + 1), iminute);
	if (!ew_CheckInteger(sHour))
		return false;
	else if (!ew_CheckRange(sHour, 0, 23))
		return false;
	if (!ew_CheckInteger(sMin))
		return false;
	else if (!ew_CheckRange(sMin, 0, 59))
		return false;
	if (iminute != -1) {
		sSec = time_value.substring(iminute + 1);		
		if (!ew_CheckInteger(sSec))
			return false;
		else if (!ew_CheckRange(sSec, 0, 59))
			return false;	
	}
	return true;
}

// Check phone
function ew_CheckPhone(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	if (object_value.length != 12)
		return false;
	if (!ew_CheckNumber(object_value.substring(0,3)))
		return false;
	else if (!ew_NumberRange((eval(object_value.substring(0,3))), 100, 1000))
		return false;
	if (object_value.charAt(3) != "-" && object_value.charAt(3) != " ")
		return false
	if (!ew_CheckNumber(object_value.substring(4,7)))
		return false;
	else if (!ew_NumberRange((eval(object_value.substring(4,7))), 100, 1000))
		return false;
	if (object_value.charAt(7) != "-" && object_value.charAt(7) != " ")
		return false;
	if (object_value.charAt(8) == "-" || object_value.charAt(8) == "+")
		return false;
	else
		return (ew_CheckInteger(object_value.substring(8,12)));
}

// Check zip
function ew_CheckZip(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	if (object_value.length != 5 && object_value.length != 10)
		return false;
	if (object_value.charAt(0) == "-" || object_value.charAt(0) == "+")
		return false;
	if (!ew_CheckInteger(object_value.substring(0,5)))
		return false;
	if (object_value.length == 5)
		return true;
	if (object_value.charAt(5) != "-" && object_value.charAt(5) != " ")
		return false;
	if (object_value.charAt(6) == "-" || object_value.charAt(6) == "+")
		return false;
	return (ew_CheckInteger(object_value.substring(6,10)));
}

// Check credit card
function ew_CheckCreditCard(object_value) {
	var white_space = " -";
	var creditcard_string = "";
	var check_char;
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	for (var i = 0; i < object_value.length; i++) {
		check_char = white_space.indexOf(object_value.charAt(i));
		if (check_char < 0)
			creditcard_string += object_value.substring(i, (i + 1));
	}	
	if (creditcard_string.length == 0)
		return false;	 
	if (creditcard_string.charAt(0) == "+")
		return false;
	if (!ew_CheckInteger(creditcard_string))
		return false;
	var doubledigit = creditcard_string.length % 2 == 1 ? false : true;
	var checkdigit = 0;
	var tempdigit;
	for (var i = 0; i < creditcard_string.length; i++) {
		tempdigit = eval(creditcard_string.charAt(i));		
		if (doubledigit) {
			tempdigit *= 2;
			checkdigit += (tempdigit % 10);			
			if ((tempdigit / 10) >= 1.0)
				checkdigit++;			
			doubledigit = false;
		}	else {
			checkdigit += tempdigit;
			doubledigit = true;
		}
	}
	return (checkdigit % 10) == 0 ? true : false;
}

// Check social security number
function ew_CheckSSC(object_value) {
	var white_space = " -+.";
	var ssc_string="";
	var check_char;
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	if (object_value.length != 11)
		return false;
	if (object_value.charAt(3) != "-" && object_value.charAt(3) != " ")
		return false;
	if (object_value.charAt(6) != "-" && object_value.charAt(6) != " ")
		return false;
	for (var i = 0; i < object_value.length; i++) {
		check_char = white_space.indexOf(object_value.charAt(i));
		if (check_char < 0)
			ssc_string += object_value.substring(i, (i + 1));
	}	
	if (ssc_string.length != 9)
		return false;	 
	if (!ew_CheckInteger(ssc_string))
		return false;
	return true;
}

// Check emails
function ew_CheckEmailList(object_value, email_cnt) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	var emailList = object_value.replace(/,/g,";");
	var arEmails = emailList.split(";");
	if (arEmails.length > email_cnt && email_cnt > 0)
		return false;
	for (var i = 0; i < arEmails.length; i++) {
		if (!ew_CheckEmail(arEmails[i]))
			return false;
	}
	return true;
}

// Check email
function ew_CheckEmail(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	object_value = object_value.replace(/^\s*|\s*$/g, "");
	var re = new RegExp("^[A-Z0-9._%+-]+@[A-Z0-9._-]+\.[A-Z]{2,6}$", "i");
	return re.test(object_value);
}

// Check GUID {xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx}	
function ew_CheckGUID(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	object_value = object_value.replace(/^\s*|\s*$/g, "");
	var re = new RegExp("^(\{{1}([0-9a-fA-F]){8}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){4}-([0-9a-fA-F]){12}\}{1})$");
	return re.test(object_value);
}

// Check file extension
function ew_CheckFileType(object_value) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	if (typeof EW_UPLOAD_ALLOWED_FILE_EXT == "undefined")
		return true;
	if (EW_UPLOAD_ALLOWED_FILE_EXT == "")
		return true;
	var fileTypes = EW_UPLOAD_ALLOWED_FILE_EXT.split(",");
	var ext = object_value.substring(object_value.lastIndexOf(".")+1, object_value.length).toLowerCase(); 
	for (var i=0; i < fileTypes.length; i++) { 
		if (fileTypes[i] == ext) 
			return true; 
	} 
	return false; 
}

// Check by regular expression
function ew_CheckByRegEx(object_value, pattern, flags) {
	if (object_value == null)
		return true;
	if (object_value.length == 0)
		return true;
	var re = (flags) ? new RegExp(pattern, flags) : new RegExp(pattern);
	return (object_value.match(re)) ? true : false;
}
