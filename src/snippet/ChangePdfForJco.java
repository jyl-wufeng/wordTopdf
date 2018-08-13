package snippet;

import java.io.File;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class ChangePdfForJco {

	public static Converter newConverter(String name) {
		return new Wps();
	}
	
	public static JacobUtil HTMLConverter(String name) {
		return new JacobUtil();
	}

	public static boolean convert(String word, String pdf) {
		// return newConverter("pdfcreator").convert(word, pdf);
		System.out.println("start1.." + word);
		return newConverter("word").convert(word, pdf);
	}
	
	public static boolean convertHTML(String word, String htmlfile) throws Exception {
		// return newConverter("pdfcreator").convert(word, pdf);
		System.out.println("start.." + word);
		HTMLConverter("word").wordToHtml2(word, htmlfile);
		return true;
	}

	public static void main(String[] args) throws InterruptedException {
		long start = System.currentTimeMillis();
		ExecutorService executorService = Executors.newFixedThreadPool(3);
		File file = new File("c:/opt/word");
		File[] files = file.listFiles();
		for (int i = 0; i < files.length; i++) {
			executorService.execute(new ConvertHTMLRunnable(files[i]));
//			new Thread(new ConvertHTMLRunnable(files[i])).start();
		}
		while (true) {
			Thread.sleep(100);
			int count = Thread.activeCount();
			if (count == 1)
				break;
		}
		long now = System.currentTimeMillis();
		String mod = String.valueOf((now - start) % 1000);
		if (mod.length() < 3)
			mod = "0" + mod;
		if (mod.length() < 3)
			mod = "0" + mod;
		System.out.println("ºÄÊ±:" + (now - start) / 1000 + "." + mod);
	}
//	public static void main(String[] args) throws InterruptedException {
//		long start = System.currentTimeMillis();
//		File file = new File("d:/opt/word");
//		File[] files = file.listFiles();
//		for (int i = 0; i < files.length; i++) {
//			new Thread(new ConvertRunnable(files[i])).start();
//		}
//		while (true) {
//			Thread.sleep(100);
//			int count = Thread.activeCount();
//			if (count == 1)
//				break;
//		}
//		long now = System.currentTimeMillis();
//		String mod = String.valueOf((now - start) % 1000);
//		if (mod.length() < 3)
//			mod = "0" + mod;
//		if (mod.length() < 3)
//			mod = "0" + mod;
//		System.out.println("è€—æ—¶:" + (now - start) / 1000 + "." + mod);
//	}
}

class ConvertRunnable implements Runnable {
	private File file;

	public ConvertRunnable(File file) {
		this.file = file;
	}

	@Override
	public void run() {
		String name = file.getName();
		DocChangePdfForJco.convert(file.getAbsolutePath(), "c:\\opt\\pdf\\" + name.substring(0, name.lastIndexOf(".")) + ".pdf");
	}

}

class ConvertHTMLRunnable implements Runnable {
	private File file;
	
	public ConvertHTMLRunnable(File file) {
		this.file = file;
	}
	
	@Override
	public void run() {
		String name = file.getName();
		try {
			System.out.println("run");
			JacobUtil.wordToHtml2(file.getAbsolutePath(), "c:\\opt\\html\\" + name.substring(0, name.lastIndexOf(".")) + ".html");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
}
