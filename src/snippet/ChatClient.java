package snippet;

import java.io.*;
import java.net.*;
import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * @author Michael Huang
 * 
 */
public class ChatClient {
	Socket s = null;
	DataOutputStream dos = null;
	DataInputStream dis = null;
	//Thread tRecv = new Thread(new RecvThread());
	private static final String UPSQL = "update cq_ywext_forpdf_pdf set stuts=?,fenjiId=?,massage=? where id=?";
    //初始化未处理文件
	private static final String INSTALSQL="update cq_ywext_forpdf_pdf set stuts='N' where stuts='E' and type<>'waterMark'";
	//查询处理文件
	private static final String SELECTSQL="select * from cq_ywext_forpdf_pdf  where stuts='N' and (type<>'waterMark' or type is null) ";
	//查询该文件是否正在处理
	private static final String ISNOW="select COUNT(*) from cq_ywext_forpdf_pdf where name=? and stuts='E'";

	//空查询避免连接断开
	private static final String ONSQL="select count(1) from cq_ywext_forpdf_pdf  ";
	public static void main(String[] args) throws FileNotFoundException {
		/*  PrintStream out = new PrintStream("E:/log/log.txt");
		  System.setOut(out);11*/
		//new ChatClient().launchFrame(Parameter.DUANKOU);
		
		Connection connection = GetConnection.GetConnection();
		ExecutorService executorService = Executors.newFixedThreadPool(3);

		try {
			PreparedStatement upps = connection.prepareStatement(INSTALSQL);
			upps.executeUpdate();
			upps.close();
			connection.setAutoCommit(false);
			while(true){
				
				try {
					boolean closed = connection.isValid(1000);
					if(closed==false){
						connection = GetConnection.GetConnection();	
						connection.setAutoCommit(false);
					}
                    //ExecutorService executorService = Executors.newFixedThreadPool(5);
                    PreparedStatement ops = connection.prepareStatement(ONSQL);
                    ops.executeQuery();
                    ops.close();
                    PreparedStatement ps = connection.prepareStatement(SELECTSQL);
                    ResultSet rs = ps.executeQuery();
                    while (rs.next()) {	
					PreparedStatement isps = connection.prepareStatement(ISNOW);
					isps.setString(1, rs.getString(2));
					ResultSet isrs = isps.executeQuery();
					String cot="0";
					while (isrs.next()) {	
						cot=isrs.getString(1);
					}
					isrs.close();
					isps.close();
					if(cot.equals("0")){
					executorService.execute(new Topdf(rs.getString(2),rs.getString(8),new JacobUtil(),rs.getString(1)));			
					   //更新状态
						PreparedStatement ps1 = connection.prepareStatement(UPSQL);
						ps1.setString(1, "E");
						ps1.setString(2, Parameter.ID);
						ps1.setString(3, null);
						ps1.setString(4, rs.getString(1));
						ps1.executeUpdate();
						ps1.close();
					
					}
                  }
                  rs.close();
                  ps.close();
                  connection.commit();
                  try {
          			Thread.sleep(500);
          		} catch (InterruptedException e1) {
          			System.out.println("延迟处理失败");
          		}
				} catch (com.mysql.jdbc.exceptions.jdbc4.CommunicationsException e) {//连接超时
					e.printStackTrace();
					connection = GetConnection.GetConnection();	
					connection.setAutoCommit(false);
				}

			
			}
		} catch (SQLException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}finally{
			try {
				connection.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

	}

/*	public void launchFrame(int port) {
		connect(port);
		tRecv.start();
	}

	public void connect(int port) {
		try {
			s = new Socket(Parameter.ZJIP, port);
			dos = new DataOutputStream(s.getOutputStream());
			dis = new DataInputStream(s.getInputStream());
			System.out.println("~~~~~~~~连接成功~~~~~~~~!");
			// 注册id
			dos.writeUTF(Parameter.ID);
			dos.flush();444444444
			bConnected = true;
		} catch (UnknownHostException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public void disconnect() {
		try {
			dos.close();
			dis.close();
			s.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	private class RecvThread implements Runnable {
		private static final String SQL = "Select * from cq_ywext_forpdf_pdf where id=? and (stuts='S' or stuts='E')";
		private static final String UPSQL = "update cq_ywext_forpdf_pdf set stuts=?,fenjiId=?,massage=? where id=?";
		private static final String DELETESQL = "delete from cq_ywext_forpdf_pdf  where id=? and type<>'redHtm'";

		String fopStu = "N";

		// 装pdf
		public void topdf(String id) {
			String name = null;
			String type = null;
			Connection connection = GetConnection.GetConnection();
			try {
				// 查询
				PreparedStatement ps1 = connection.prepareStatement(SQL);
				ps1.setString(1, id);
				ResultSet rs = ps1.executeQuery();
				while (rs.next()) {
					name = rs.getString(2);
					type = rs.getString(8);
				}
				rs.close();
				ps1.close();
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}

			// 处理
			if(name!=null&&type!=null){
			Topdf topdf = new Topdf(name, type);
			fopStu = topdf.startT();
			System.out.println("成功返回");
			try {
				if (fopStu.equals("Y")) {
					if(!type.equals("redHtm")){
					PreparedStatement ps = connection.prepareStatement(DELETESQL);
					ps.setString(1, id);
					PreparedStatement ps = connection.prepareStatement(UPSQL);
					ps.setString(1, "Y");
					ps.setString(2, Parameter.ID);
					ps.setString(3, null);
					ps.setString(4, id);
					 
					ps.executeUpdate();
					ps.close();}else{
						PreparedStatement ps = connection.prepareStatement(DELETESQL);
						ps.setString(1, id);
						PreparedStatement ps = connection.prepareStatement(UPSQL);
						ps.setString(1, "Y");
						ps.setString(2, Parameter.ID);
						ps.setString(3, null);
						ps.setString(4, id);
						 
						ps.executeUpdate();
						ps.close();
					}
				} else {
					topdf(id);
				}
			
			} catch (SQLException e) {
				// TODO Auto-generated catch block
				GetConnection.close(connection);
				e.printStackTrace();
			}
			}
			GetConnection.close(connection);
		}

		public void run() {
			try {
				while (bConnected) {
					String id = dis.readUTF();
					// taContent.setText(taContent.getText() + id + '\n');
					try {
						topdf(id);
						dos.writeUTF("Y");
						dos.flush();
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
			} catch (SocketException e) {
				System.out.println("退出了，bye!");
			} catch (EOFException e) {
				System.out.println("退出了，bye!");
			} catch (IOException e) {
				e.printStackTrace();
			}

		}

	}*/

}