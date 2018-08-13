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
		  //���
		  FileOutputStream out = null;
		  //PDFҳ��
		  int pageNum = 0;
		  //PDF����
		  int x,y = 0;
		  //PDF���ƶ���
		  Dispatch pdfObject = null;
		  //PDF�������
		  Dispatch pointxy = null;
		  //pdfActiveX PDDoc���� ��Ҫ����PDF����
		  ActiveXComponent app = new ActiveXComponent("AcroExch.PDDoc");
		  //pdfActiveX PDF���������
		  ActiveXComponent point = new ActiveXComponent("AcroExch.Point");
		  try {
		  //�õ����ƶ���
		  pdfObject = app.getObject();
		  //�õ��������
		  pointxy = point.getObject();
		  //��PDF�ļ�������PDF�����Ŀ�ʼ
		  Dispatch.call(pdfObject, "Open", new Variant(filepath));
		  //�õ���ǰ��PDF�ļ���ҳ��
		  pageNum = Dispatch.call(pdfObject, "GetNumPages").toInt(); 
		  
		  for(int i=0;i < pageNum;i++){
			  //����ҳ��õ���ҳPDF
			  Dispatch page = Dispatch.call(pdfObject, "AcquirePage", new Variant(i)).toDispatch();
			  //�õ�PDF��ҳ��С��Point����
			  Dispatch pagePoint = Dispatch.call(page, "GetSize").toDispatch();
			  //����PDFλ�ö���Ϊ����ͼƬ����������׼��
			  ActiveXComponent pdfRect = new ActiveXComponent("AcroExch.Rect");
			  //�õ���ҳPDF�Ŀ�
			  int imgWidth = (int) (Dispatch.get(pagePoint, "x").toInt() * 2);
			  //�õ���ҳPDF�ĸ�
            int imgHeight = (int) (Dispatch.get(pagePoint, "y").toInt() * 2);
            //����PDFλ�ö���
            Dispatch pdfRectDoc = pdfRect.getObject();
            //����PDFλ�ö����ֵ
            Dispatch.put(pdfRectDoc, "Left", new Integer(0));
            Dispatch.put(pdfRectDoc, "Right", new Integer(imgWidth));
            Dispatch.put(pdfRectDoc, "Top", new Integer(0));
            Dispatch.put(pdfRectDoc, "Bottom", new Integer(imgHeight));
            //�����ú�λ�õ�PDF������Windows���а壬������λ�ö���,����㣬����㣬�ֱ���
            Dispatch.call(page, "CopyToClipboard",new Object[]{pdfRectDoc,0,0,200});
            Image image = getImageFromClipboard();
            BufferedImage tag = new BufferedImage(imgWidth, imgHeight, 8);
            Graphics graphics = tag.getGraphics();
            graphics.drawImage(image, 0, 0, null);
            graphics.dispose();
            //���ͼƬ
            ImageIO.write(tag, "JPEG", new File(savePath+i+".jpg"));
		}
			} catch (Exception e) {
				e.printStackTrace();
			} 
		  finally {
				//�ر�PDF
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
