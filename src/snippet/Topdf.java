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
/*
wordתpdf������
 */
public class Topdf implements Runnable{
	//����״̬
	private static final String UPSQL = "update jyl_word_processing set STATUS=?,IDENTIFICATION=?,EXTEND1=? where ID=?";

	private String status = "N";
	private String srcPath;
	private String desPath;
	private String type;
	private JacobUtil jacobUtil;
	private String ID;

	//����������ʼ������ֵ
	public Topdf(String srcPath,String desPath, String type,JacobUtil jacobUtil,String id) {
		this.srcPath = srcPath;
		this.desPath=desPath;
		this.type=type;
		this.status = "N";
		this.ID=id;
		this.jacobUtil=jacobUtil;
	}

	public void run() {
		System.out.println("����" + srcPath);

		//·�������ڴ���·��
		File path = new File(desPath.substring(0,
				desPath.lastIndexOf("/")));
		if (!path.isDirectory()) {
			path.mkdir();
		}

		final ExecutorService exec = Executors.newFixedThreadPool(1);
		Callable<String> call = new Callable<String>() {
			public String call() throws Exception {
				String message="";
				// ��ʼִ�к�ʱ����
				srcPath = srcPath.replaceAll(Parameter.PREFIX,
						Parameter.SRCPREFIX);
				// תpdf
				if (type.equals("pdf")) {
					File srcfile = new File(srcPath);
					File desfile = new File(desPath);
					if (!srcfile.isFile()) {
						message = "�ļ�������";
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "-----" + "�ļ�������");
						status = "C";
						return message;
					}
					if (desfile.isFile()) {
						desfile.delete();
					}

					if (jacobUtil.convert(srcPath, desPath)) {
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "-----" + "�ɹ�");
						status = "Y";
					} else {
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "-----" + "ʧ��");
						status = "N";
					}
				}// תhtml
				else {
					File srcfile = new File(srcPath);
					File desfile = new File(desPath);
					if (!srcfile.isFile()) {
						message = "�ļ�������";
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "----" + "�ļ�������");
						status="C";
						return message;
					}
					if (desfile.isFile()) {
						desfile.delete();
					}
					boolean wordToHtml2 = jacobUtil.wordToHtml2(srcPath, desPath);
					if (wordToHtml2) {
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "-----" + "�ɹ�");
						status = "Y";
					} else {
						System.out.println(new Date(System.currentTimeMillis())
								+ "�������ļ�-" + srcPath + "          " + desPath
								+ "-----" + "ʧ��");
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
		Connection connection = DbcpUtil.getConnection();

		if(status.equals("Y")){//����ɹ�
			updateStatus(connection,"Y");
		}else if(status.equals("C")){  //�ļ�������
			updateStatus(connection, "C");
		} else{ //����ʧ�����´���
			try {
				invoke();
				updateStatus(connection, "E");
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		try {
			connection.commit();
			connection.close();
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("�������");
	}

	private void updateStatus(Connection connection, String c)   {
		try {
		PreparedStatement ps = connection.prepareStatement(UPSQL);
		ps.setString(1, c);
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

	// �ر�word
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
