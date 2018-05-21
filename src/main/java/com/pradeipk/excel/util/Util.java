package com.pradeipk.excel.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class Util {
	
	
	public static String getDate(Date date) {
		SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd");
		String formattedDate = formatter.format(date);
		System.out.println(formattedDate);
		/*StringBuilder new_date = new StringBuilder(date.getDate());
		new_date.append("-");
		new_date.append(date.getMonth());
		new_date.append("-");
		new_date.append(date.get);
		new_date.append("-");*/
		return formattedDate;
		
	}

}
