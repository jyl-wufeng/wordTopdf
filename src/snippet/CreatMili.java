package snippet;


import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;

public class CreatMili {
public static void main(String[] args){
	  File file = new File("F:/xcd/�½��ı��ĵ�.txt");
      BufferedReader reader = null;
      try {
          reader = new BufferedReader(new FileReader(file));
          String tempString = null;
          int line = 1;
          // һ�ζ���һ�У�ֱ������nullΪ�ļ�����
          while ((tempString = reader.readLine()) != null) {
              // ��ʾ�к�
        	  Mosaic(tempString);
          //    tempString.lastIndexOf(".");
              line++;
          }
          reader.close();
      } catch (IOException e) {
          e.printStackTrace();
      } finally {
          if (reader != null) {
              try {
                  reader.close();
              } catch (IOException e1) {
              }
          }
      }
}

public static void Mosaic(String str){
	String substring = str.substring(str.lastIndexOf(".")+1, str.length());
	if(substring.equals("java")){
		String[] split = str.split("/");
		String mil;
		if(split[8].equals("ywext")){
		 mil=	"jar -uf ../WEB-INF/lib/ywext-java-0.1-SNAPSHOT.jar  ."+
					str.substring(str.lastIndexOf("java/")+4, str.length()).replaceAll(".java", ".class");
          
		}else{
			 mil=	"jar -uf ../WEB-INF/lib/portal-java-0.1-SNAPSHOT.jar  ."+
					str.substring(str.lastIndexOf("java/")+4, str.length()).replaceAll(".java", ".class");
		}
		System.out.println(mil);

	}
	if(substring.equals("jsp")){
		String mil="mv "+str.substring(str.lastIndexOf("/")+1, str.length())
				+" ."+str.substring(str.lastIndexOf("webapp/")+6, str.length());
		String mil1="jar -uf portal.war"+" ."+str.substring(str.lastIndexOf("webapp/")+6, str.length());
		System.out.println(mil);	
		System.out.println(mil1);	

	}
	if(substring.equals("xml")){
		String[] split = str.split("/");
		String mil;
		if(split[8].equals("ywext")){
		 mil=	"jar -uf ../WEB-INF/lib/ywext-java-0.1-SNAPSHOT.jar  ."+
				str.substring(str.lastIndexOf("java/")+4, str.length());
          
		}else{
			 mil=	"jar -uf ../WEB-INF/lib/portal-java-0.1-SNAPSHOT.jar  ."+
					str.substring(str.lastIndexOf("java/")+4, str.length());
		}
		System.out.println(mil);
	}
}
}
