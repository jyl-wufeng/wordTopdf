package snippet;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Properties;  
import java.util.Vector;
import org.slf4j.Logger;  
import org.slf4j.LoggerFactory;  
import com.ibm.icu.text.SimpleDateFormat;
import com.jcraft.jsch.Channel;  
import com.jcraft.jsch.ChannelSftp;  
import com.jcraft.jsch.ChannelSftp.LsEntry;
import com.jcraft.jsch.JSch;  
import com.jcraft.jsch.Session;  


/**
 * 
 * </strong> <strong>Description : </strong>SCM附件同步工具类<br>
 * <strong>Create on : 2017年12月18日<br>
 * </strong> <li>
 * <strong>Copyright (C) Mocha Software Co.,Ltd.<br>
 * </strong> <li>
 * 
 * @author 王文杰 wangwj@mochasoft.com.cn<br>
 * @version <strong>v1.0</strong><br>
 * <br>
 *          <strong>修改历史:</strong><br>
 *          修改人 修改日期 修改描述<br>
 *          -------------------------------------------<br>
 * <br>
 * <br>
 */


public class SFTPUtilForSSH {
	private static final Logger LOG = LoggerFactory.getLogger(SFTPUtilForSSH.class);  
	
	private ChannelSftp sftp;  
	private Channel channel;  
	private Session sshSession;  
	
	/**
	 * 登录sftp
	 * @param host 目标IP
	 * @param port 目标端口
	 * @param username 登录账号
	 * @param password 登录密码
	 * @return
	 */
    public boolean login(String host, int port, String username, final String password) {  
        try {
        	LOG.info("开始登录sftp........");
            JSch jsch = new JSch();
            jsch.getSession(username, host, port);
            sshSession = jsch.getSession(username, host, port);
            sshSession.setPassword(password);
            Properties sshConfig = new Properties();
            sshConfig.put("StrictHostKeyChecking", "no");
            sshSession.setConfig(sshConfig);
            sshSession.connect();  
            LOG.info("Session connected!"+"SFTP连接成功，用户名====="+username);
            channel = sshSession.openChannel("sftp");  
            channel.connect();  
            LOG.debug("Channel connected!");  
            this.sftp = (ChannelSftp) channel;  
        } catch (Exception e) { 
            e.printStackTrace();  
            LOG.info(username + "登录SFTP服务失败！" + e.getMessage()); 
            return false;
        } 
        return true;  
    }  
  
    public void closeChannel(Channel channel) {  
        if (channel != null) {  
            if (channel.isConnected()) {  
                channel.disconnect();  
            }  
        }  
    }  
  
    public void closeSession(Session session) {  
        if (session != null) {  
            if (session.isConnected()) {  
                session.disconnect();  
            }  
        }  
    } 
    
    public void logout(){
    	closeChannel(sftp);  
        closeChannel(channel);  
        closeSession(sshSession);  
    }
    
    
    /*** 
     * 上传文件 
     * @param localFile 本地文件 
     * @param romotUpLoadePath  上传服务器路径 - 应该以/结束 
     * @param fileName 要上传的文件名
     * @param sftpPath sftp服务器的文件存放根目录，服务端提供- 应该以/结束 
     * */  
    @SuppressWarnings("unchecked")
	public boolean uploadFile(File localFile, String romotUpLoadePath,String fileName,String sftpPath,String prjId) {  
    	FileInputStream fileInStream = null;  
        boolean success = false;  
        try {
        	
        	// 判断路径   yyyyMM  是否存在
			Vector<LsEntry> vector = this.sftp.ls(sftpPath);
        	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        	// 0 标识目录不存在 1表示目录存在
        	String xm = "0";
        	for(LsEntry entry : vector){
        		// 判断是否存在 yyyyMM 路径
        		if(sdf.format(new Date()).equals(entry.getFilename())){
        			xm = "1";
        			break;
        		}
        	}
        	
        	// 不存在 yyyyMM 路径 则创建路径 sftpPath+sdf.format(new Date())+"/"
        	if(xm.equals("0")){
        		sftp.mkdir(sftpPath+sdf.format(new Date())+"/");
        	}
        	// 判断是否存在 prjId  路径
        	String mk = "0";
        	Vector<LsEntry> vectors = this.sftp.ls(sftpPath+sdf.format(new Date())+"/");
        	for(LsEntry entrys : vectors){
        		if(prjId.equals(entrys.getFilename())){
        			mk = "1";
        			break;
        		}
        	}
        	
        	// 不存在路径则创建
        	if(mk.equals("0")){
        		sftp.mkdir(romotUpLoadePath);
        	}
        	this.sftp.cd(romotUpLoadePath);
            LOG.info(localFile.getName() + "开始上传.....");  
            
            this.sftp.put(new FileInputStream(localFile), fileName);
           
            LOG.info(localFile.getName() + "上传成功");
        	
			success=true;
			LOG.info(localFile.getName() + "上传成功"); 
            return success; 
		} catch (Exception e) {
			e.printStackTrace();
		}finally {  
            if (fileInStream != null) {  
                try {  
                	fileInStream.close();  
                } catch (IOException e) {  
                    e.printStackTrace();  
                }  
            }  
        }
        return success;  
    } 
    
    
    /*** 
     * @上传文件夹
     * @param localDirectory  当地文件全路径  例：/xxx/xxx/xxx.txt
     * @param remoteDirectoryPath 
     *            Ftp 服务器路径 以目录"/"结束 
     * *//*  
	public boolean uploadDirectory(String localDirectory,
			String remoteDirectoryPath) {
		File src = new File(localDirectory);

		remoteDirectoryPath = remoteDirectoryPath + src.getName() + "/";
		File remoteFile = new File(remoteDirectoryPath);
		if (!remoteFile.exists()) {
			remoteFile.mkdir();
		}

		File[] allFile = src.listFiles();
		for (int currentFile = 0; currentFile < allFile.length; currentFile++) {
			if (!allFile[currentFile].isDirectory()) {
				String srcName = allFile[currentFile].getPath().toString();
				uploadFile(new File(srcName), remoteDirectoryPath,
						allFile[currentFile].getName());
			}
		}
		for (int currentFile = 0; currentFile < allFile.length; currentFile++) {
			if (allFile[currentFile].isDirectory()) {
				// 递归
				uploadDirectory(allFile[currentFile].getPath().toString(),
						remoteDirectoryPath);
			}
		}
		return true;
	}*/
}
