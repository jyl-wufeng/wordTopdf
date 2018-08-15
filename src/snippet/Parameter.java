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
	public static String PREFIX;
	public static String ID;
	public static String PATH;
	public static  String InitSize;
	/*
	 * 
	 * 初始化属性
	 * 
	 * 
	 * */
	static {
		Properties prop = new Properties();
       // System.out.println(Thread.currentThread().getContextClassLoader().getResource("parameter.properties"));
	//	System.out.println(Parameter.class.getResource("/"));
		PATH=Parameter.class.getResource("").toString().replace("file:/","")+"parameter.properties";
		InputStream in = null;
		try {
			in = new BufferedInputStream(new FileInputStream(
					PATH));
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
		PREFIX= prop.getProperty("PREFIX");
		ID= prop.getProperty("ID");
		SRCPREFIX= prop.getProperty("SRCPREFIX");
		InitSize=prop.getProperty("InitSize");
		try {
			in.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}
}
