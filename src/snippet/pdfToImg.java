package snippet;

import java.awt.Graphics;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;

import javax.imageio.ImageIO;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

public class pdfToImg {
	public static void main(String[] args){
		savaPageAsJpgByAcrobat("H:\\moban\\ww1.pdf","H:\\moban\\ww1.jpg");
	}
	public static void savaPageAsJpgByAcrobat(String filepath,String savePath){
		  //输出
		  FileOutputStream out = null;
		  //PDF页数
		  int pageNum = 0;
		  //PDF宽、高
		  int x,y = 0;
		  //PDF控制对象
		  Dispatch pdfObject = null;
		  //PDF坐标对象
		  Dispatch pointxy = null;
		  //pdfActiveX PDDoc对象 主要建立PDF对象
		  ActiveXComponent app = new ActiveXComponent("AcroExch.PDDoc");
		  //pdfActiveX PDF的坐标对象
		  ActiveXComponent point = new ActiveXComponent("AcroExch.Point");
		  try {
		  //得到控制对象
		  pdfObject = app.getObject();
		  //得到坐标对象
		  pointxy = point.getObject();
		  //打开PDF文件，建立PDF操作的开始
		  Dispatch.call(pdfObject, "Open", new Variant(filepath));
		  //得到当前打开PDF文件的页数
		  pageNum = Dispatch.call(pdfObject, "GetNumPages").toInt(); 
		  
		  for(int i=0;i < pageNum;i++){
			  //根据页码得到单页PDF
			  Dispatch page = Dispatch.call(pdfObject, "AcquirePage", new Variant(i)).toDispatch();
			  //得到PDF单页大小的Point对象
			  Dispatch pagePoint = Dispatch.call(page, "GetSize").toDispatch();
			  //创建PDF位置对象，为拷贝图片到剪贴板做准备
			  ActiveXComponent pdfRect = new ActiveXComponent("AcroExch.Rect");
			  //得到单页PDF的宽
			  int imgWidth = (int) (Dispatch.get(pagePoint, "x").toInt() * 2);
			  //得到单页PDF的高
            int imgHeight = (int) (Dispatch.get(pagePoint, "y").toInt() * 2);
            //控制PDF位置对象
            Dispatch pdfRectDoc = pdfRect.getObject();
            //设置PDF位置对象的值
            Dispatch.put(pdfRectDoc, "Left", new Integer(0));
            Dispatch.put(pdfRectDoc, "Right", new Integer(imgWidth));
            Dispatch.put(pdfRectDoc, "Top", new Integer(0));
            Dispatch.put(pdfRectDoc, "Bottom", new Integer(imgHeight));
            //将设置好位置的PDF拷贝到Windows剪切板，参数：位置对象,宽起点，高起点，分辨率
            Dispatch.call(page, "CopyToClipboard",new Object[]{pdfRectDoc,0,0,200});
            Image image = getImageFromClipboard();
            BufferedImage tag = new BufferedImage(imgWidth, imgHeight, 8);
            Graphics graphics = tag.getGraphics();
            graphics.drawImage(image, 0, 0, null);
            graphics.dispose();
            //输出图片
            ImageIO.write(tag, "JPEG", new File(savePath+i+".jpg"));
		}
			} catch (Exception e) {
				e.printStackTrace();
			} 
		  finally {
				//关闭PDF
				 app.invoke("Close", new Variant[] {}); 
			}
		 
	}
	
	public static Image getImageFromClipboard() throws Exception {
		Clipboard sysc = Toolkit.getDefaultToolkit().getSystemClipboard();
		Transferable cc = sysc.getContents(null);
		if (cc == null)
			return null;
		else if (cc.isDataFlavorSupported(DataFlavor.imageFlavor))
			return (Image) cc.getTransferData(DataFlavor.imageFlavor);
		return null;
	}
}
