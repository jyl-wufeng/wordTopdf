package snippet;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileInputStream;
import java.io.BufferedInputStream;
import java.util.Iterator;
import java.util.Properties;

public class Parameter {
	public static String DRIVER;
	public static String PASSWROD;
	public static String USERNAME;
	public static String URL;
	public static String SRCPREFIX;
	public static String DECPREFIX;
	public static String ZJIP;
	public static int DUANKOU;
	public static String ID;
	
	/*
	 * 
	 * 初始化属性
	 * 
	 * 
	 * */
	static {
		Properties prop = new Properties();

		InputStream in = null;
		try {
			in = new BufferedInputStream(new FileInputStream(
					"C:/pdfserver/parameter.properties"));
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		}
		// 读取属性文件a.properties
		try {
			prop.load(in);
		} catch (IOException e) {
			e.printStackTrace();
		} // /加载属性列表
		DRIVER = prop.getProperty("DRIVER");
		PASSWROD = prop.getProperty("PASSWROD");
		USERNAME = prop.getProperty("USERNAME");
		URL = prop.getProperty("URL");
		SRCPREFIX = prop.getProperty("SRCPREFIX");
		DECPREFIX = prop.getProperty("DECPREFIX");
		ZJIP = prop.getProperty("ZJIP");
		DUANKOU = Integer.valueOf(prop.getProperty("DUANKOU"));
		ID = prop.getProperty("ID");
		try {
			in.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
