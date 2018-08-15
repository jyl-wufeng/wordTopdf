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

	// 8/10 ����word�����html
	public static final int WORD_HTML = 10;
	private static final int wdFormatPDF = 17;
	//word����
	public  ActiveXComponent app=null;
    //word���ĵ�
    public  Dispatch  doc=null;

public static List<String> list;
	/**
	 * WORDתHTML
	 * 
	 * @param docfile
	 *            WORD�ļ�ȫ·��
	 * @param htmlfile
	 *            ת����HTML���·��
	 */
	public boolean wordToHtml(String docfile, String htmlfile) {
		// ����wordӦ�ó���(Microsoft Office Word 2007)
		try {
		 app = new ActiveXComponent("Word.Application");
		// ����wordӦ�ó��򲻿ɼ�
		app.setProperty("Visible", new Variant(false));
		// documents��ʾword����������ĵ����ڣ���word�Ƕ��ĵ�Ӧ�ó���
		Dispatch docs = app.getProperty("Documents").toDispatch();
		// ��Ҫת����word�ļ�
		 doc = Dispatch.call(docs, "Open", docfile, false, true)
				.toDispatch();
		Dispatch selection = Dispatch.get(app, "Selection").toDispatch();
		// ���ñ��� 65001----utf-8
		Dispatch option = Dispatch.get(doc, "WebOptions").toDispatch();
		Dispatch.put(option, "Encoding", 65001);
		// �����޶�
		Dispatch.call(doc, "AcceptAllRevisions");
		/*if(htmlfile.indexOf("redHtm")==-1){*/
		// ���б��
		Dispatch tables = Dispatch.get(doc, "Tables").toDispatch();
		// ��ȡ�����Ŀ
		int tablesCount = Dispatch.get(tables, "Count").toInt();
		// ѭ����ȡ���
		for (int i = 1; i <= tablesCount; i++) {
			// ��ȡ��i�����
			Dispatch table = Dispatch.call(tables, "Item", new Variant(i))
					.toDispatch();
			Dispatch rows = Dispatch.call(table, "Rows").toDispatch();
			// ���ñ�񲻿ɻ���
			Dispatch.put(rows, "WrapAroundText", new Variant(false));
			//���ñ���Զ��ص��ߴ�����Ӧ����
			//Dispatch.put(table, "AllowAutoFit", new Variant(true));	
			// ���ñ�����
			// Dispatch Range = Dispatch.call(table, "Range").toDispatch();
			// Dispatch ParagraphFormat = Dispatch.call(Range,
			// "ParagraphFormat").toDispatch(); ��������ݾ���1���� 2����
			
			//�������0
			//Dispatch.put(rows, "Alignment", 2);	
			Dispatch.put(rows, "LeftIndent",0.0);
			Double sad = Dispatch.get(rows, "LeftIndent").toDouble();
			Dispatch.put(rows, "LeftIndent",0.0);
			//������
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
		// ���htm
		Dispatch.call(doc, "SaveAs", htmlfile, WORD_HTML);
		Dispatch.call(doc, "Close", false);
		doc=null;
		app.invoke("Quit", 0);
		app=null;
		boolean mvimage=mvImage(htmlfile);
		if(!mvimage){
			System.out.println("ͼƬ����"+htmlfile+"---����");
			return false;
		}
		return true;
		} catch (Exception e) {
			return false;
		}
	}
	/**
	 * ���������ַ���   �������ֱ����
	 * 
	 * @param doc
	 *            WORD�ļ�ȫ·��
	 * @param tableIndex
	 *            �ڼ������
     * @param cellRowIdx
	 *            �ڼ���      
     * @param cellColIdx
	 *            �ڼ���               
	 */
	 public static String getTxtFromCell(Dispatch selection,Dispatch doc,int tableIndex, int cellRowIdx, int cellColIdx) {
		//  System.out.println(tableIndex+cellRowIdx+cellColIdx);
		 // ���б��
		  Dispatch tables = Dispatch.get(doc, "Tables").toDispatch(); 
		  // Ҫ���ı��
		  Dispatch table = Dispatch.call(tables, "Item", new Variant(tableIndex)).toDispatch();	  
		  Dispatch cell = Dispatch.call(table, "Cell", new Variant(cellRowIdx),new Variant(cellColIdx)).toDispatch();
		  Dispatch Range=Dispatch.get(cell,"Range").toDispatch();
		  String text=Dispatch.get(Range,"Text").toString();
		  text = text.substring(0, text.length() - 2); // ȥ�����Ļس���;
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
	 * EXCELתPDF
	 * 
	 * @param inFilePath
	 *            WORD�ļ�ȫ·��
	 * @param outFilePath
	 *            ת����PDF���·��
	 */
	public static boolean xlsToPdf(String inFilePath, String outFilePath) {
		ActiveXComponent ax=null;
		//��ʼ��com���߳�  
        ComThread.InitSTA();  
		try {
		    ComThread.InitSTA(true);
		    ax = new ActiveXComponent("Excel.Application");
			ax.setProperty("Visible", new Variant(false));
			ax.setProperty("AutomationSecurity", new Variant(3)); // ���ú�
			Dispatch excels = ax.getProperty("Workbooks").toDispatch();

			Dispatch excel = Dispatch.invoke(
					excels,
					"Open",
					Dispatch.Method,
					new Object[] { inFilePath, new Variant(false),
							new Variant(false) }, new int[9]).toDispatch();
			// ת����ʽ
			Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method,
					new Object[] { new Variant(0), // PDF��ʽ=0
							outFilePath, new Variant(0) // 0=��׼ (���ɵ�PDFͼƬ�����ģ��)
														// 1=��С�ļ�
														// (���ɵ�PDFͼƬ����һ����Ϳ)
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
				// �ر�wordӦ�ó���
				if(ax!=null){
					ax.invoke("Quit", new Variant[] {});}
				//�ر�com���߳�  
				ComThread.Release();
			} catch (Exception e) {
				System.out.println("�ر��쳣");

			}  
	}
	}
	//htmlͼƬ����
	public static  boolean mvImage(String htmPath){
		try {
			String SQL="SELECT e.PATH FROM mocha_document_file_resource e,mocha_document_file t WHERE e.ID = t.RESOURCE_ID and t.BOINS_ID =? AND e.`NAME` =?";
			String htmPathr=htmPath.replace(".htm", ".files/");
			String insId=htmPath.substring(htmPath.lastIndexOf("/")+1, htmPath.lastIndexOf("."));
			//��ѯ�Ƿ�����ݿ�
			 Connection connection = DbcpUtil.getConnection();
			File file=new File(htmPathr);
			 File[] files = file.listFiles();
			 String imgPath = null;
			 if(files!=null&&files.length>0){
			 for(File file1:files){
				String imageName= file1.getName() ;
			     imgPath=null;

				try {
					// ��ѯ
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
				//���浽���ݿ�
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
				
			 //�ƶ�ͼƬ
			boolean mvimage=copyFolder(htmPathr,imgPath);
			if(!mvimage){
				System.out.println("�ƶ�ͼƬ��"+htmPathr+"----"+imgPath+"---����");
				return false;
			}
			 }}
		} catch (Exception e) {
			return false;
		}
		return true;
	}
	
	/** 
	* ���������ļ������� 
	* @param oldPath String ԭ�ļ�·�� �磺c:/fqf 
	* @param newPath String ���ƺ�·�� �磺f:/fqf/ff 
	* @return boolean 
	*/ 
	public static  boolean copyFolder(String oldPath, String newPath) { 

	try { 
	(new File(newPath)).mkdirs(); //����ļ��в����� �������ļ��� 
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
	if(temp.isDirectory()){//��������ļ��� 
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
     * @param inputFile  Դ�ļ�·��
     * @param outputFile Ŀ���ļ�·��
     * @param waterMarkName  ˮӡ��������
     */
    public static void addWaterMarkIncludeWords(String imageFile, String inputFile, String outputFile, String waterMarkName ) {  
        try {  
            PdfReader reader = new PdfReader(inputFile);  
            PdfStamper stamper = new PdfStamper(reader, new FileOutputStream( outputFile));  
            // BaseFont base = BaseFont.createFont(BaseFont.COURIER_BOLD , BaseFont.CP1250 , BaseFont.NOT_EMBEDDED);  
            Image image = Image.getInstance(imageFile);
            /*BaseFont base =BaseFont.createFont();*/
            BaseFont base = BaseFont.createFont("STSong-Light",
            		"UniGB-UCS2-H", BaseFont.NOT_EMBEDDED); // ���Ĵ���
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
                // ����͸����Ϊ0.2
                gs.setFillOpacity(0.6f);
                under.setGState(gs);
                // ע��������ֺ�������һ��restoreState ����������Ч          	
                // ��ʼ
                under.beginText();
            	// ������ɫ Ĭ��Ϊ��ɫ
                under.setColorFill(BaseColor.GRAY);
            	// �������弰�ֺ�
                under.setFontAndSize(base, 30);
            	// ������ʼλ��
                under.setTextMatrix(100, 800);    
                float width = pageRect.getWidth();
                float height = pageRect.getHeight();
                under.showTextAligned(Element.ALIGN_LEFT, waterMarkName,(width/2)-142,(height/2)-126, 50);
                //under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, 250,290, 50);
                under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, (width/2)-52,(height/2)-126, 50);
               // under.addImage(image);
                //͸��������
                // ע������������һ��restoreState ����������Ч          	
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
     * @param path  Դ�ļ�·��
   
     */
    public  boolean creatPdf(String path){
    	//wordתpdf
    	path=path.replaceAll("\\\\", "/");
    	String[] split = path.split("\\*");
		String srcPath= split[0];
		String decPath= split[1];
		File filepath=new File(decPath.substring(0, decPath.lastIndexOf("/")));
		if(!filepath.isDirectory()){
			filepath.mkdirs();
		}
    	if(path.indexOf("doc")!=-1||path.indexOf(".docx")!=-1){
    		 //��ʼ��com���߳�  
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
    
/*	// ������ɫ Ĭ��Ϊ��ɫ
	under.setColorFill(BaseColor.GRAY);
	// �������弰�ֺ�
    under.setFontAndSize(base, 25);
	// ������ʼλ��
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
	// ������ɫ Ĭ��Ϊ��ɫ
	under.setColorFill(BaseColor.GRAY);
	// �������弰�ֺ�
	under.setFontAndSize(base, 50);
	// ������ʼλ��
	under.setTextMatrix(100, 800);
	under.showTextAligned(Element.ALIGN_LEFT, waterMarkName, height,
			240, 45);*/
    
    public static void aaa() throws InterruptedException{
        Thread.sleep(1000 * 5);  

    }
	
    /**  
     * WORDתHTML  
     * @param docfile WORD�ļ�ȫ·��  
     * @param htmlfile ת����HTML���·��  
     */  
    public static boolean wordToHtml2(String docfile, String htmlfile)   
    {   
    	// ����wordӦ�ó���(Microsoft Office Word 2003)
        ActiveXComponent app = new ActiveXComponent("Word.Application");
        System.out.println("*****ת��...*****" + Thread.currentThread().getName());
        try  
        {	
        	// ����wordӦ�ó��򲻿ɼ�  
            app.setProperty("Visible", new Variant(false));  
            // documents��ʾword����������ĵ����ڣ���word�Ƕ��ĵ�Ӧ�ó���
            Dispatch docs = app.getProperty("Documents").toDispatch();  
            // ��Ҫת����word�ļ�
            Dispatch doc = Dispatch.invoke(   
                    docs,   
                    "Open",   
                    Dispatch.Method,   
                    new Object[] { docfile, new Variant(false), 
                    		new Variant(true) }, new int[1]).toDispatch();
         // ���ñ��� 65001----utf-8
    		Dispatch option = Dispatch.get(doc, "WebOptions").toDispatch();
    		Dispatch.put(option, "Encoding", 65001);
    		// �����޶�
    		Dispatch.call(doc, "AcceptAllRevisions");
            // ��Ϊhtml��ʽ���浽��ʱ�ļ�
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] {   
                    htmlfile, new Variant(WORD_HTML) }, new int[1]);   
            // �ر�word�ļ�
            Dispatch.call(doc, "Close", new Variant(false));   
        }   
        catch (Exception e)   
        {   
            e.printStackTrace();   
            return false;
        }   
        finally  
        {   
        	//�ر�wordӦ�ó���
            app.invoke("Quit", new Variant[] {});   
        } 
        System.out.println("*****ת�����********" + htmlfile);
        boolean mvimage=mvImage(htmlfile);
		if(!mvimage){
			System.out.println("ͼƬ����"+htmlfile+"---����");
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
        // �����޶�WordBasic
       // Dispatch WordBasic = Dispatch.get(doc, "WordBasic").toDispatch();
		  Dispatch.call(WordBasic, "AcceptAllChangesInDoc");
		  //����
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