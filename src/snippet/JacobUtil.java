package snippet;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

import com.itextpdf.text.Element;
import com.itextpdf.text.Image;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfGState;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfStamper;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.itextpdf.text.BaseColor;

public class JacobUtil {

	// 8/10 代表word保存成html
	public static final int WORD_HTML = 10;
	private static final int wdFormatPDF = 17;
	//word进程
	public  ActiveXComponent app=null;
    //word打开文档
    public  Dispatch  doc=null;

public static List<String> list;
	/**
	 * WORD转HTML
	 * 
	 * @param docfile
	 *            WORD文件全路径
	 * @param htmlfile
	 *            转换后HTML存放路径
	 */
	public boolean wordToHtml(String docfile, String htmlfile) {
		// 启动word应用程序(Microsoft Office Word 2007)
		try {
		 app = new ActiveXComponent("Word.Application");
		// 设置word应用程序不可见
		app.setProperty("Visible", new Variant(false));
		// documents表示word程序的所有文档窗口，（word是多文档应用程序）
		Dispatch docs = app.getProperty("Documents").toDispatch();
		// 打开要转换的word文件
		 doc = Dispatch.call(docs, "Open", docfile, false, true)
				.toDispatch();
		Dispatch selection = Dispatch.get(app, "Selection").toDispatch();
		// 设置编码 65001----utf-8
		Dispatch option = Dispatch.get(doc, "WebOptions").toDispatch();
		Dispatch.put(option, "Encoding", 65001);
		// 接受修订
		Dispatch.call(doc, "AcceptAllRevisions");
		/*if(htmlfile.indexOf("redHtm")==-1){*/
		// 所有表格
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// 获取表格数目
		int tablesCount = Dispatch.get(tables, "Count").toInt();
		// 循环获取表格
		for (int i = 1; i <= tablesCount; i++) {
			// 获取第i个表格
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i))
					.toDispatch();
			Dispatch rows = Dispatch.call(table, "Rows").toDispatch();
			// 设置表格不可环绕
			Dispatch.put(rows, "WrapAroundText", new Variant(false));
			//设置表格自动重调尺寸以适应内容
			//Dispatch.put(table, "AllowAutoFit", new Variant(true));	
			// 设置表格居中
			// Dispatch Range = Dispatch.call(table, "Range").toDispatch();
			// Dispatch ParagraphFormat = Dispatch.call(Range,
			// "ParagraphFormat").toDispatch(); 表格里内容居中1居中 2居右
			
			//表格缩进0
			//Dispatch.put(rows, "Alignment", 2);	
			Dispatch.put(rows, "LeftIndent",0.0);
			Double sad = Dispatch.get(rows, "LeftIndent").toDouble();
			Dispatch.put(rows, "LeftIndent",0.0);
			//表格居中
			Dispatch.put(rows, "Alignment", 1);			  
			Dispatch columns = Dispatch.call(table, "Columns").toDispatch();
			int rowsCount = Dispatch.get(rows, "Count").toInt();
			int columnsCount = Dispatch.get(columns, "Count").toInt();
			for(int j=1;j<=rowsCount;j++){
			 for(int k=1;k<=columnsCount;k++){
				 try {
					getTxtFromCell(selection,doc,i,j,k);
				} catch (Exception e) {
					
				}
			 }
			}
		}
		Dispatch.call(doc, "AcceptAllRevisions");
		// 另存htm
		Dispatch.call(doc, "SaveAs", htmlfile, WORD_HTML);
		Dispatch.call(doc, "Close", false);
		doc=null;
		app.invoke("Quit", 0);
		app=null;
		boolean mvimage=mvImage(htmlfile);
		if(!mvimage){
			System.out.println("图片处理："+htmlfile+"---出错");
			return false;
		}
		return true;
		} catch (Exception e) {
			return false;
		}
	}
	/**
	 * 处理表格文字方向   竖排文字变横排
	 * 
	 * @param doc
	 *            WORD文件全路径
	 * @param tableIndex
	 *            第几个表格
     * @param cellRowIdx
	 *            第几行      
     * @param cellColIdx
	 *            第几列               
	 */
	 public static String getTxtFromCell(Dispatch selection,Dispatch doc,int tableIndex, int cellRowIdx, int cellColIdx) {
		//  System.out.println(tableIndex+cellRowIdx+cellColIdx);
		 // 所有表格
		  Dispatch tables = Dispatch.get(doc, "Tables").toDispatch(); 
		  // 要填充的表格
		  Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();	  
		  Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),new Variant(cellColIdx)).toDispatch();
		  Dispatch Range=Dispatch.get(cell,"Range").toDispatch();
		  String text=Dispatch.get(Range,"Text").toString();
		  text = text.substring(0, text.length() - 2); // 去掉最后的回车符;
		  char[] chars=text.toCharArray();
		  String ret="";
		  Dispatch.call(cell, "Select");
		 // Dispatch ret = Dispatch.get(selection, "Text").toDispatch();
		 ret= Dispatch.get(selection, "Orientation").toString();
		 if(ret.equals("9999999")){
			 Dispatch.put(selection, "Orientation", 0);
			 String str="";
			 for(char a:chars){
				str=str+a+"\n" ;
			 }
			 str = str.substring(0, str.length() - 1);
			 Dispatch.put(selection, "Text", str);
			 System.out.println("---------");
		 }
		  return ret;
		 }
	/**
	 * EXCEL转PDF
	 * 
	 * @param inFilePath
	 *            WORD文件全路径
	 * @param outFilePath
	 *            转换后PDF存放路径
	 */
	public static boolean xlsToPdf(String inFilePath, String outFilePath) {
		ActiveXComponent ax=null;
		//初始化com的线程  
        ComThread.InitSTA();  
		try {
		    ComThread.InitSTA(true);
		    ax = new ActiveXComponent("Excel.Application");
			ax.setProperty("Visible", new Variant(false));
			ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();

			Dispatch excel = Dispatch.invoke(
					excels,
					"Open",
					Dispatch.Method,
					new Object[] { inFilePath, new Variant(false),
							new Variant(false) }, new int[9]).toDispatch();
			// 转换格式
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method,
					new Object[] { new Variant(0), // PDF格式=0
							outFilePath, new Variant(0) // 0=标准 (生成的PDF图片不会变模糊)
														// 1=最小文件
														// (生成的PDF图片糊的一塌糊涂)
					}, new int[1]);

			Dispatch.call(excel, "Close", new Variant(false));

			if (ax != null) {
				ax.invoke("Quit", new Variant[] {});
				ax = null;
			}
			ComThread.Release();
			return true;
		} catch (Exception es) {
			return false;
		}finally {
			try {
				// 关闭word应用程序
				if(ax!=null){
					ax.invoke("Quit", new Variant[] {});}
				//关闭com的线程  
				ComThread.Release();
			} catch (Exception e) {
				System.out.println("关闭异常");

			}  
	}
	}
	//html图片处理
	public static  boolean mvImage(String htmPath){
		try {
			String SQL="SELECT e.PATH FROM mocha_document_file_resource e,mocha_document_file t WHERE e.ID = t.RESOURCE_ID and t.BOINS_ID =? AND e.`NAME` =?";
			String htmPathr=htmPath.replace(".htm", ".files/");
			String insId=htmPath.substring(htmPath.lastIndexOf("/")+1, htmPath.lastIndexOf("."));
			//查询是否存数据库
			 Connection connection = DbcpUtil.getConnection();
			File file=new File(htmPathr);
			 File[] files = file.listFiles();
			 String imgPath = null;
			 if(files!=null&&files.length>0){
			 for(File file1:files){
				String imageName= file1.getName() ;
			     imgPath=null;

				try {
					// 查询
					PreparedStatement ps1 = connection.prepareStatement(SQL);
					ps1.setString(1, insId);
					ps1.setString(2, imageName);
					ResultSet rs = ps1.executeQuery();
					while (rs.next()) {
						imgPath = rs.getString(1);
					}
					rs.close();
					ps1.close();
				} catch (SQLException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				System.out.println(imgPath);
				//保存到数据库
				if(imgPath==null){
					imgPath=htmPath.substring(htmPath.indexOf("htm/")+4,htmPath.lastIndexOf("/"))+"/"+insId+"/"+imageName;	
					System.out.println(imgPath+"---");
					try {
						String uuid1="docres-"+UUID.randomUUID().toString().toUpperCase();
						String uuid2="doc-"+UUID.randomUUID().toString().toUpperCase();
						PreparedStatement up2 = connection.prepareStatement("insert into mocha_document_file_resource(id,name,type,path)values(?,?,?,?)");
						up2.setString(1, uuid1);
						up2.setString(2, imageName);
						up2.setString(3, "img");
						up2.setString(4, imgPath);
						up2.executeUpdate();
						up2.close();
						PreparedStatement up1 = connection.prepareStatement("insert into mocha_document_file(document_file_id,boins_id,document_file_type,resource_id,extend1) values(?,?,?,?,?)");
						up1.setString(1, uuid2);
						up1.setString(2, insId);
						up1.setString(3, "img");
						up1.setString(4, uuid1);
						up1.setString(5, imageName);
						up1.executeUpdate();
						up1.close();
					} catch (SQLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			   
			 }
			 connection.close();
			 if(imgPath!=null){
			 imgPath=htmPathr.substring(0,htmPathr.indexOf("gw/")+3)
					 +"img/"
					 +imgPath.substring(0, imgPath.lastIndexOf("/")+1);
				
			 //移动图片
			boolean mvimage=copyFolder(htmPathr,imgPath);
			if(!mvimage){
				System.out.println("移动图片："+htmPathr+"----"+imgPath+"---出错");
				return false;
			}
			 }}
		} catch (Exception e) {
			return false;
		}
		return true;
	}
	
	/** 
	* 复制整个文件夹内容 
	* @param oldPath String 原文件路径 如：c:/fqf 
	* @param newPath String 复制后路径 如：f:/fqf/ff 
	* @return boolean 
	*/ 
	public static  boolean copyFolder(String oldPath, String newPath) { 

	try { 
	(new File(newPath)).mkdirs(); //如果文件夹不存在 则建立新文件夹 
	File a=new File(oldPath); 
	String[] file=a.list(); 
	File temp=null; 
	for (int i = 0; i < file.length; i++) { 
	if(oldPath.endsWith(File.separator)){ 
	temp=new File(oldPath+file[i]); 
	} 
	else{ 
	temp=new File(oldPath+File.separator+file[i]); 
	} 

	if(temp.isFile()){ 
	FileInputStream input = new FileInputStream(temp); 
	FileOutputStream output = new FileOutputStream(newPath + "/" + 
	(temp.getName()).toString()); 
	byte[] b = new byte[1024 * 5]; 
	int len; 
	while ( (len = input.read(b)) != -1) { 
	output.write(b, 0, len); 
	} 
	output.flush(); 
	output.close(); 
	input.close(); 
	} 
	if(temp.isDirectory()){//如果是子文件夹 
	copyFolder(oldPath+"/"+file[i],newPath+"/"+file[i]); 
	} 
	}
	return true;
	} 
	catch (Exception e) { 
    return false;
	} 
	}
	
	 /**
     * @param inputFile  源文件路径
     * @param outputFile 目标文件路径
     * @param waterMarkName  水印文字内容
     */
    public static void addWaterMarkIncludeWords(String imageFile, String inputFile, String outputFile, String waterMarkName ) {  
        try {  
            PdfReader reader = new PdfReader(inputFile);  
            PdfStamper stamper = new PdfStamper(reader, new FileOutputStream( outputFile));  
            // BaseFont base = BaseFont.createFont(BaseFont.COURIER_BOLD , BaseFont.CP1250 , BaseFont.NOT_EMBEDDED);  
            Image image = Image.getInstance(imageFile);
            /*BaseFont base =BaseFont.createFont();*/
            BaseFont base = BaseFont.createFont("STSong-Light",
            		"UniGB-UCS2-H", BaseFont.NOT_EMBEDDED); // 中文处理
            int total = reader.getNumberOfPages() + 1;  
           // System.out.println(stamper.g);
            PdfContentByte under;   
            Rectangle pageRect = null;
            for (int i = 1; i < total; i++) {  
                //under = stamper.getUnderContent(i); 
            /*	stamper.set*/
            	  pageRect = stamper.getReader().
                          getPageSizeWithRotation(i);
            	 // under=stamper.getUnderContent(i);
            	 // under.
                under=stamper.getOverContent(i);
                under.saveState();
                PdfGState gs = new PdfGState();
                // 设置透明度为0.2
                gs.setFillOpacity(0.6f);
                under.setGState(gs);
                // 注意添加文字后必须调用一次restoreState 否则设置无效          	
                // 开始
                under.beginText();
            	// 设置颜色 默认为蓝色
                under.setColorFill(BaseColor.GRAY);
            	// 设置字体及字号
                under.setFontAndSize(base, 30);
            	// 设置起始位置
                under.setTextMatrix(100, 800);    
                float width = pageRect.getWidth();
                float height = pageRect.getHeight();
                under.showTextAligned(Element.ALIGN_LEFT, waterMarkName,(width/2)-142,(height/2)-126, 50);
                //under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 250,290, 50);
                under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, (width/2)-52,(height/2)-126, 50);
               // under.addImage(image);
                //透明度设置
                // 注意这里必须调用一次restoreState 否则设置无效          	
               // under.restoreState();  
                under.endText();
                under.setLineWidth(1f);
                under.stroke();
            }  
            stamper.close();  
        } catch (Exception e) {  
            e.printStackTrace();  
        }  
    }  
    /**
     * @param path  源文件路径
   
     */
    public  boolean creatPdf(String path){
    	//word转pdf
    	path=path.replaceAll("\\\\", "/");
    	String[] split = path.split("\\*");
		String srcPath= split[0];
		String decPath= split[1];
		File filepath=new File(decPath.substring(0, decPath.lastIndexOf("/")));
		if(!filepath.isDirectory()){
			filepath.mkdirs();
		}
    	if(path.indexOf("doc")!=-1||path.indexOf(".docx")!=-1){
    		 //初始化com的线程  
    		  try {
                 app = new ActiveXComponent("Word.Application");  
                 Dispatch docs = app.getProperty("Documents").toDispatch();  
                  doc = Dispatch.call(docs, "Open", srcPath, false, true)  
                         .toDispatch();  
                 Dispatch.call(doc, "ExportAsFixedFormat", decPath, 17);
                 Dispatch.call(doc, "Close", false);  
                 doc=null;
                 app.invoke("Quit", 0);  
                 app=null;
                 return true;  
             } catch (Exception e) {  
                 return false;  
             }
    	}else if(path.indexOf("xls")!=-1||path.indexOf(".xlsx")!=-1){
    		return xlsToPdf(srcPath,decPath);
    	}
		return true;
    }
    
/*	// 设置颜色 默认为蓝色
	under.setColorFill(BaseColor.GRAY);
	// 设置字体及字号
    under.setFontAndSize(base, 25);
	// 设置起始位置
    under.setTextMatrix(100, 800);
   // under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 110,240, 45);
    //under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 130,240, 45);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 10,20, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 189,20, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 380,20, 50);
    
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 10,310, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 189,310, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 380,310, 50);
    
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 10,600, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 189,600, 50);
    under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 380,600, 50);*/
    
    
/*    
	// 设置颜色 默认为蓝色
	under.setColorFill(BaseColor.GRAY);
	// 设置字体及字号
	under.setFontAndSize(base, 50);
	// 设置起始位置
	under.setTextMatrix(100, 800);
	under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, height,
			240, 45);*/
    
    public static void aaa() throws InterruptedException{
        Thread.sleep(1000 * 5);  

    }
	
    /**  
     * WORD转HTML  
     * @param docfile WORD文件全路径  
     * @param htmlfile 转换后HTML存放路径  
     */  
    public static boolean wordToHtml2(String docfile, String htmlfile)   
    {   
    	// 启动word应用程序(Microsoft Office Word 2003)
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        System.out.println("*****转换...*****" + Thread.currentThread().getName());
        try  
        {	
        	// 设置word应用程序不可见  
            app.setProperty("Visible", new Variant(false));  
            // documents表示word程序的所有文档窗口，（word是多文档应用程序）
            Dispatch docs = app.getProperty("Documents").toDispatch();  
            // 打开要转换的word文件
            Dispatch doc = Dispatch.invoke(   
                    docs,   
                    "Open",   
                    Dispatch.Method,   
                    new Object[] { docfile, new Variant(false), 
                    		new Variant(true) }, new int[1]).toDispatch();
         // 设置编码 65001----utf-8
    		Dispatch option = Dispatch.get(doc, "WebOptions").toDispatch();
    		Dispatch.put(option, "Encoding", 65001);
    		// 接受修订
    		Dispatch.call(doc, "AcceptAllRevisions");
            // 作为html格式保存到临时文件
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] {   
                    htmlfile, new Variant(WORD_HTML) }, new int[1]);   
            // 关闭word文件
            Dispatch.call(doc, "Close", new Variant(false));   
        }   
        catch (Exception e)   
        {   
            e.printStackTrace();   
            return false;
        }   
        finally  
        {   
        	//关闭word应用程序
            app.invoke("Quit", new Variant[] {});   
        } 
        System.out.println("*****转换完毕********" + htmlfile);
        boolean mvimage=mvImage(htmlfile);
		if(!mvimage){
			System.out.println("图片处理："+htmlfile+"---出错");
			return false;
		}
        return true;
    }
    
    public   boolean convert(String word, String pdf) {
	      try{
			  System.out.println("bb"+pdf);
        app = new ActiveXComponent("Word.Application");
			  System.out.println("aa"+pdf);
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
	  e.printStackTrace();
  	  System.out.println(e.toString());
  	  return false;  
    }
	  
	    }


	    public static void main(String[] args){
			String property = System.getProperty("java.library.path");
			System.out.println(property);
			JacobUtil jo=new JacobUtil();
			jo.convert("H:/moban/345.doc","H:/moban/pdf/345.pdf");
		}
}