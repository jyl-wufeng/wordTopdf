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
public class TopdfStart {
	Socket s = null;
	DataOutputStream dos = null;
	DataInputStream dis = null;
	//Thread tRecv = new Thread(new RecvThread());
	private static final String UPSQL = "update jyl_word_processing set STATUS=?,IDENTIFICATION=?,EXTEND1=? where ID=?";
    //初始化未处理文件
	//private static final String INSTALSQL="update cq_ywext_forpdf_pdf set stuts='E' where stuts<>'S'" ;
	//查询处理文件
	private static final String SELECTSQL="select * from jyl_word_processing  where STATUS='E' order by date";
	//查询该文件是否正在处理
	private static final String ISNOW="select COUNT(*) from jyl_word_processing where SRCPATH=? and STATUS='S'";

	//空查询避免连接断开
	private static final String ONSQL="select count(1) from jyl_word_processing  ";
	public static void main(String[] args) throws FileNotFoundException {
		Connection connection = DbcpUtil.getConnection();
		//初始化线程池数量 5
		ExecutorService executorService = Executors.newFixedThreadPool(3);

		try {
		/*	PreparedStatement upps = connection.prepareStatement(INSTALSQL);
			upps.executeUpdate();
			upps.close();*/
			//设置不自动提交
			connection.setAutoCommit(false);
			while(true){
				
				try {
					//-----判断连接是否有效----
					boolean closed = connection.isValid(1000);
					if(closed==false){
						connection = DbcpUtil.getConnection();
						connection.setAutoCommit(false);
					}
                    //ExecutorService executorService = Executors.newFixedThreadPool(5);
                    PreparedStatement ops = connection.prepareStatement(ONSQL);
                    ops.executeQuery();
                    ops.close();
                    PreparedStatement ps = connection.prepareStatement(SELECTSQL);
                    ResultSet rs = ps.executeQuery();
                    while (rs.next()) {
                    	//判断该文件是否正在处理
					PreparedStatement isps = connection.prepareStatement(ISNOW);
					isps.setString(1, rs.getString(2));
					ResultSet isrs = isps.executeQuery();
					String cot="0";
					while (isrs.next()) {	
						cot=isrs.getString(1);
					}
					isrs.close();
					isps.close();
					//未被处理状态为0
					if(cot.equals("0")){
                        //更新状态
                        PreparedStatement ps1 = connection.prepareStatement(UPSQL);
                        ps1.setString(1, "S");
                        ps1.setString(2, Parameter.ID);
                        ps1.setString(3, null);
                        ps1.setString(4, rs.getString(1));
                        ps1.executeUpdate();
                        ps1.close();
					executorService.execute(new Topdf(rs.getString(2),rs.getString(3),rs.getString(7),new JacobUtil(),rs.getString(1)));
					}
                  }
                  rs.close();
                  ps.close();
                  connection.commit();
                  //扫描间隔
                  Thread.sleep(500);

				} catch (Exception e) {//连接超时
					e.printStackTrace();
					connection = DbcpUtil.getConnection();
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

}