package snippet;

import java.io.File;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Date;
import java.util.concurrent.Callable;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
public class Topdf implements Runnable{
	private static final String SQL = "Select * from cq_ywext_forpdf_pdf where id=? and (stuts='S' or stuts='E')";
	private static final String DELETESQL = "delete from cq_ywext_forpdf_pdf  where id=? and (type<>'redHtm' or type is null)";
	private static final String UPSQL = "update cq_ywext_forpdf_pdf set stuts=?,fenjiId=?,massage=? where id=?";

	public String status = "N";
	String srcPath1;
	String type;
	JacobUtil jacobUtil;
	String ID;
	public Topdf(String srcPath, String type1,JacobUtil jacobUtil,String id) {
		this.srcPath1 = srcPath;
		if(type1!=null&&!type1.equals("null")){
		this.type = type1;}else{
			this.type="pdf"	;
		}
		this.status = "N";
		this.ID=id;
		this.jacobUtil=jacobUtil;
	}

	public void run() {
		try {
			Thread.sleep(1000);
		} catch (InterruptedException e1) {
			System.out.println("延迟处理失败");
		}
		System.out.println("接收" + srcPath1);
		final ExecutorService exec = Executors.newFixedThreadPool(1);
		Callable<String> call = new Callable<String>() {
			public String call() throws Exception {
				// 开始执行耗时操作

				srcPath1 = srcPath1.replaceAll("/app/sharedata1",
						Parameter.SRCPREFIX);
				// 转pdf
				if (type.equals("pdf")) {
					/*
					 * String srcPath=Parameter.SRCPREFIX+srcPath1;
					 * System.out.print(srcPath1+"---------------"); String
					 * desPath
					 * =Parameter.SRCPREFIX+srcPath1.substring(0,srcPath1.
					 * indexOf
					 * ("redWord")-1)+"/pdf"+srcPath1.substring(srcPath1.lastIndexOf
					 * ("/"),srcPath1.length()).replaceAll("doc", "pdf"); String
					 * message ; srcPath=srcPath.replaceAll("/sharedata", "");
					 * desPath=desPath.replaceAll("/sharedata", "");
					 */
					String message;
					String srcPath = srcPath1;
					String desPath = srcPath1.substring(0,
							srcPath1.indexOf("redWord") - 1)
							+ "/pdf"
							+ srcPath1.substring(srcPath1.lastIndexOf("/"),
									srcPath1.length()).replaceAll("doc", "pdf");
					File srcfile = new File(srcPath);
					File desfile = new File(desPath);
					if (!srcfile.isFile()) {
						message = "文件不存在";
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "-----" + "文件不存在");
						//return message;
					}
					if (desfile.isFile()) {
						desfile.delete();
					}
					if (jacobUtil.convert(srcPath, desPath)) {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "-----" + "成功");
						status = "Y";
					} else {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "-----" + "失败");
						status = "N";
					}
				}
				// 转水印pdf
				else if (type.equals("waterMark")) {
					System.out.println(srcPath1);
				
					boolean sta = jacobUtil.creatPdf(srcPath1);
					if (sta) {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath1 + "-----" + "成功");
						status = "Y";
					} else {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath1 + "-----" + "失败");
						status = "N";
					}
				}// 转html
				else {

					String srcPath = srcPath1;
					System.out.print(srcPath1 + "---------------");
					// Parameter.SRCPREFIX+srcPath1.substring(0,srcPath1.indexOf("redWord")-1)+"/pdf"+srcPath1.substring(srcPath1.lastIndexOf("/"),srcPath1.length()).replaceAll("doc",
					// "htm");
					String desPath;
					String message;

					/*
					 * String message; String srcPath=srcPath1; String desPath;
					 */
					if (type.equals("htm")) {
						desPath = srcPath1.replaceAll("word", type);
					} else {
						desPath = srcPath1.replaceAll("redWord", type);
					}
					desPath = desPath.substring(0, desPath.lastIndexOf("doc"))
							+ "htm";
					File path = new File(desPath.substring(0,
							desPath.lastIndexOf("/")));
					if (!path.isDirectory()) {
						path.mkdir();
					}
					File srcfile = new File(srcPath);
					File desfile = new File(desPath);
					if (!srcfile.isFile()) {
						message = "文件不存在";
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "----" + "文件不存在");
						//return message;
					}
					if (desfile.isFile()) {
						desfile.delete();
					}
					boolean wordToHtml2 = jacobUtil.wordToHtml2(srcPath, desPath);
					if (wordToHtml2) {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "-----" + "成功");
						status = "Y";
					} else {
						System.out.println(new Date(System.currentTimeMillis())
								+ "：处理文件-" + srcPath + "          " + desPath
								+ "-----" + "失败");
						status = "N";
					}

				}
				return "";
			}
		};
		java.util.concurrent.Future<String> future = exec.submit(call);
		int runDate = 1;
		boolean isout = false;
		while (runDate < 50 && !isout) {
			try {
				Thread.sleep(2000);
			} catch (InterruptedException e) {
				e.printStackTrace();
			}
			isout = future.isDone();
			System.out.println(runDate);
			runDate++;
		}
		//invoke();
		Connection connection = GetConnection.GetConnection();

		if(status.equals("Y")){
			
			if(!type.equals("redHtm")){
				try {
					PreparedStatement ps = connection.prepareStatement(DELETESQL);
					ps.setString(1, ID);
					/*PreparedStatement ps = connection.prepareStatement(UPSQL);
					ps.setString(1, "Y");
					ps.setString(2, Parameter.ID);
					ps.setString(3, null);
					ps.setString(4, id);*/
					 
					ps.executeUpdate();
					ps.close();
				} catch (SQLException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				}else{
					try {
						PreparedStatement ps = connection.prepareStatement(DELETESQL);
						ps.setString(1, ID);
						/*PreparedStatement ps = connection.prepareStatement(UPSQL);
						ps.setString(1, "Y");
						ps.setString(2, Parameter.ID);
						ps.setString(3, null);
						ps.setString(4, ID);*/
						 
						ps.executeUpdate();
						ps.close();
					} catch (SQLException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
		}else{
			try {
				invoke();
			} catch (Exception e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
			
			//更新数据库状态
			try {
				PreparedStatement ps = connection.prepareStatement(UPSQL);
				ps.setString(1, "N");
				ps.setString(2, Parameter.ID);
				ps.setString(3, null);
				ps.setString(4, ID);
				ps.executeUpdate();
				ps.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		try {
			connection.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("处理完成");
		// Topdf.invoke();

	}

	// 关闭word
	public static void invoke() {
		Runtime run = Runtime.getRuntime();
		try {
			// run.exec("cmd /k shutdown -s -t 3600");
			Process process = run.exec("cmd.exe /c tskill winword");
			InputStream in = process.getInputStream();
			while (in.read() != -1) {
				System.out.println(in.read());
			}
			in.close();
			process.waitFor();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
