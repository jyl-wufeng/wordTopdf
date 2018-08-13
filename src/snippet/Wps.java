package snippet;


import com.jacob.activeX.ActiveXComponent;  
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;  
import com.jacob.com.Variant;


public  class Wps implements Converter {  
   private static final int wdFormatPDF = 17; 
	//word进程
public  ActiveXComponent app=null;
//word打开文档
public  Dispatch  doc=null;
/*   public static void main(String[] args){
	   String word="F:/sharedata/document/gw/redWord/2017/09/00-DCOEWN6D-BDWV-WAFJ-X5VO-Y5S18H9EMN1M.doc";
	   String pdf="F:/sharedata/document/gw/redWord/2017/09/00-DCOEWN6D-BDWV-WAFJ-X5VO-Y5S18H9EMN1M.pdf";
	   ActiveXComponent app = new ActiveXComponent("Word.Application");
		// 设置word应用程序不可见
		//app.setProperty("Visible", new Variant(false));
		// documents表示word程序的所有文档窗口，（word是多文档应用程序）
		Dispatch docs = app.getProperty("Documents").toDispatch();
		// 打开要转换的word文件
		Dispatch doc = Dispatch.call(docs, "Open", word, false, true)
				.toDispatch();
		Dispatch.call(doc, "SaveAs", pdf, wdFormatPDF);
		Dispatch.call(doc, "Close", false);

       app.invoke("Quit", 0);
   }*/
    public   boolean convert(String word, String pdf) {
    	      try{
              app = new ActiveXComponent("Word.Application");  
              Dispatch docs = app.getProperty("Documents").toDispatch();
              Dispatch WordBasic = app.getProperty("WordBasic").toDispatch();  
              doc = Dispatch.call(docs, "Open", word, false, false)  
                      .toDispatch();  
              // 接受修订WordBasic
             // Dispatch WordBasic = Dispatch.get(doc, "WordBasic").toDispatch();
      		  Dispatch.call(WordBasic, "AcceptAllChangesInDoc");
      		  //保存
      		  Dispatch.call(doc, "Save");
      	//	Dispatch.call(doc,"Close",new Variant(true));
              Dispatch.call(doc, "ExportAsFixedFormat", pdf, wdFormatPDF);
              Dispatch.call(doc, "Close", false);  
              doc=null;
              app.invoke("Quit", new Variant[] {});
              app=null;
              return true;  
          } catch (Exception e) {  
        	  System.out.println(e.toString());
        	  return false;  
          }
    	  
    	    }  
    
}  
