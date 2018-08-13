package snippet;

public class DocChangePdfForJco {  
  
    public static Converter newConverter(String name) {  
      return new Wps();  }
  
  
    public synchronized static boolean convert(String word, String pdf) {  
    	return newConverter("word").convert(word, pdf);  
    }  
}  

