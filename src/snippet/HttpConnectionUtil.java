package snippet;


import java.io.BufferedInputStream;
import java.io.BufferedReader;
import java.io.DataOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.net.HttpURLConnection;
import java.net.InetSocketAddress;
import java.net.MalformedURLException;
import java.net.Proxy;
import java.net.URL;
import java.net.URLConnection;
import java.util.Iterator;
import java.util.Map;

/**
 * Javaԭ����API�����ڷ���HTTP���󣬼�java.net.URL��java.net.URLConnection����ЩAPI�ܺ��á��ܳ��ã�
 * ��������㣻
 * 
 * 1.ͨ��ͳһ��Դ��λ����java.net.URL����ȡ��������java.net.URLConnection�� 2.��������Ĳ��� 3.��������
 * 4.������������ʽ��ȡ�������� 5.�ر�������
 * 
 * @author H__D
 *
 */
public class HttpConnectionUtil {


    /**
     * ���ļ��ϴ��ķ���
     * 
     * @param actionUrl���ϴ���·��
     * @param uploadFilePaths����Ҫ�ϴ����ļ�·��������
     * @return
     */
    @SuppressWarnings("finally")
    public static String uploadFile(String actionUrl, String[] uploadFilePaths) {
        String end = "\r\n";
        String twoHyphens = "--";
        String boundary = "*****";

        DataOutputStream ds = null;
        InputStream inputStream = null;
        InputStreamReader inputStreamReader = null;
        BufferedReader reader = null;
        StringBuffer resultBuffer = new StringBuffer();
        String tempLine = null;

        try {
            // ͳһ��Դ
            URL url = new URL(actionUrl);
            Proxy proxy = new Proxy(java.net.Proxy.Type.HTTP,  
            new InetSocketAddress("127.0.0.1", 8888)); 
            // ������ĸ��࣬������
            URLConnection urlConnection = url.openConnection(proxy);
            // http��������
            HttpURLConnection httpURLConnection = (HttpURLConnection) urlConnection;

            // �����Ƿ��httpUrlConnection���룬Ĭ���������true;
            httpURLConnection.setDoInput(true);
            // �����Ƿ���httpUrlConnection���
            httpURLConnection.setDoOutput(true);
            // Post ������ʹ�û���
            httpURLConnection.setUseCaches(false);
            // �趨����ķ�����Ĭ����GET
            httpURLConnection.setRequestMethod("POST");
            // �����ַ��������Ӳ���
            httpURLConnection.setRequestProperty("Connection", "Keep-Alive");
            // �����ַ�����
            httpURLConnection.setRequestProperty("Charset", "UTF-8");
            // ����������������
            httpURLConnection.setRequestProperty("Content-Type", "multipart/form-data;boundary=" + boundary);

            // ����DataOutputStream
            ds = new DataOutputStream(httpURLConnection.getOutputStream());
            for (int i = 0; i < uploadFilePaths.length; i++) {
                String uploadFile = uploadFilePaths[i];
                String filename = uploadFile.substring(uploadFile.lastIndexOf("//") + 1);
                ds.writeBytes(twoHyphens + boundary + end);
                ds.writeBytes("Content-Disposition: form-data; " + "name=\"file" + i + "\";filename=\"" + filename
                        + "\"" + end);
                ds.writeBytes(end);
                FileInputStream fStream = new FileInputStream(uploadFile);
                int bufferSize = 1024;
                byte[] buffer = new byte[bufferSize];
                int length = -1;
                while ((length = fStream.read(buffer)) != -1) {
                    ds.write(buffer, 0, length);
                }
                ds.writeBytes(end);
                /* close streams */
                fStream.close();
            }
            ds.writeBytes(twoHyphens + boundary + twoHyphens + end);
            /* close streams */
           // ds.write("WordFlag=htm&BoInsId=00-VYV2MZ2Z-OXFJ-AYZ9-0JTP-ZGGHKMAWRL41".getBytes("UTF-8"));
            ds.flush();
            if (httpURLConnection.getResponseCode() >= 300) {
                throw new Exception(
                        "HTTP Request is not success, Response code is " + httpURLConnection.getResponseCode());
            }

            if (httpURLConnection.getResponseCode() == HttpURLConnection.HTTP_OK) {
                inputStream = httpURLConnection.getInputStream();
                inputStreamReader = new InputStreamReader(inputStream);
                reader = new BufferedReader(inputStreamReader);
                tempLine = null;
                resultBuffer = new StringBuffer();
                while ((tempLine = reader.readLine()) != null) {
                    resultBuffer.append(tempLine);
                    resultBuffer.append("\n");
                }
            }

        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally {
            if (ds != null) {
                try {
                    ds.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
            if (reader != null) {
                try {
                    reader.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
            if (inputStreamReader != null) {
                try {
                    inputStreamReader.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }
            if (inputStream != null) {
                try {
                    inputStream.close();
                } catch (IOException e) {
                    // TODO Auto-generated catch block
                    e.printStackTrace();
                }
            }

            return resultBuffer.toString();
        }
    }


    public static void main(String[] args) {

        // �ϴ��ļ�����
         String str = uploadFile("http://10.189.88.207:8080/portal-web/bpm/document/save?WordFlag=htm&BoInsId=00-VYV2MZ2Z-OXFJ-AYZ9-0JTP-ZGGHKMAWRL41",new String[] { "H:/moban/aa.rar" });
         System.out.println(str);


    }

}

