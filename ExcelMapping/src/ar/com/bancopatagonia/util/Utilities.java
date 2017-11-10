package ar.com.bancopatagonia.util;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Utilities {
	
	public static String dataString() {
		DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
		Date date = new Date();
		String datestring = dateFormat.format(date);
		return datestring;
	}

}
