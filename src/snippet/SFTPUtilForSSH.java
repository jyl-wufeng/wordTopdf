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
 * </strong> <strong>Description : </strong>SCM����ͬ��������<br>
 * <strong>Create on : 2017��12��18��<br>
 * </strong> <li>
 * <strong>Copyright (C) Mocha Software Co.,Ltd.<br>
 * </strong> <li>
 * 
 * @author ���Ľ� wangwj@mochasoft.com.cn<br>
 * @version <strong>v1.0</strong><br>
 * <br>
 *          <strong>�޸���ʷ:</strong><br>
 *          �޸��� �޸����� �޸�����<br>
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
	 * ��¼sftp
	 * @param host Ŀ��IP
	 * @param port Ŀ��˿�
	 * @param username ��¼�˺�
	 * @param password ��¼����
	 * @return
	 */
    public boolean login(String host, int port, String username, final String password) {  
        try {
        	LOG.info("��ʼ��¼sftp........");
            JSch jsch = new JSch();
            jsch.getSession(username, host, port);
            sshSession = jsch.getSession(username, host, port);
            sshSession.setPassword(password);
            Properties sshConfig = new Properties();
            sshConfig.put("StrictHostKeyChecking", "no");
            sshSession.setConfig(sshConfig);
            sshSession.connect();  
            LOG.info("Session connected!"+"SFTP���ӳɹ����û���====="+username);
            channel = sshSession.openChannel("sftp");  
            channel.connect();  
            LOG.debug("Channel connected!");  
            this.sftp = (ChannelSftp) channel;  
        } catch (Exception e) { 
            e.printStackTrace();  
            LOG.info(username + "��¼SFTP����ʧ�ܣ�" + e.getMessage()); 
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
     * �ϴ��ļ� 
     * @param localFile �����ļ� 
     * @param romotUpLoadePath  �ϴ�������·�� - Ӧ����/���� 
     * @param fileName Ҫ�ϴ����ļ���
     * @param sftpPath sftp���������ļ���Ÿ�Ŀ¼��������ṩ- Ӧ����/���� 
     * */  
    @SuppressWarnings("unchecked")
	public boolean uploadFile(File localFile, String romotUpLoadePath,String fileName,String sftpPath,String prjId) {  
    	FileInputStream fileInStream = null;  
        boolean success = false;  
        try {
        	
        	// �ж�·��   yyyyMM  �Ƿ����
			Vector<LsEntry> vector = this.sftp.ls(sftpPath);
        	SimpleDateFormat sdf = new SimpleDateFormat("yyyyMM");
        	// 0 ��ʶĿ¼������ 1��ʾĿ¼����
        	String xm = "0";
        	for(LsEntry entry : vector){
        		// �ж��Ƿ���� yyyyMM ·��
        		if(sdf.format(new Date()).equals(entry.getFilename())){
        			xm = "1";
        			break;
        		}
        	}
        	
        	// ������ yyyyMM ·�� �򴴽�·�� sftpPath+sdf.format(new Date())+"/"
        	if(xm.equals("0")){
        		sftp.mkdir(sftpPath+sdf.format(new Date())+"/");
        	}
        	// �ж��Ƿ���� prjId  ·��
        	String mk = "0";
        	Vector<LsEntry> vectors = this.sftp.ls(sftpPath+sdf.format(new Date())+"/");
        	for(LsEntry entrys : vectors){
        		if(prjId.equals(entrys.getFilename())){
        			mk = "1";
        			break;
        		}
        	}
        	
        	// ������·���򴴽�
        	if(mk.equals("0")){
        		sftp.mkdir(romotUpLoadePath);
        	}
        	this.sftp.cd(romotUpLoadePath);
            LOG.info(localFile.getName() + "��ʼ�ϴ�.....");  
            
            this.sftp.put(new FileInputStream(localFile), fileName);
           
            LOG.info(localFile.getName() + "�ϴ��ɹ�");
        	
			success=true;
			LOG.info(localFile.getName() + "�ϴ��ɹ�"); 
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
     * @�ϴ��ļ���
     * @param localDirectory  �����ļ�ȫ·��  ����/xxx/xxx/xxx.txt
     * @param remoteDirectoryPath 
     *            Ftp ������·�� ��Ŀ¼"/"���� 
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
				// �ݹ�
				uploadDirectory(allFile[currentFile].getPath().toString(),
						remoteDirectoryPath);
			}
		}
		return true;
	}*/
}
