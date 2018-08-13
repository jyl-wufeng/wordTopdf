package snippet;

import java.util.concurrent.Callable;
import java.util.concurrent.ExecutionException;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.TimeoutException;



public class ThreadTest {  
  
    public static void main(String[] args) throws InterruptedException,  
            ExecutionException {  
          
        final ExecutorService exec = Executors.newFixedThreadPool(1);  
          
        Callable<String> call = new Callable<String>() {  
            public String call() throws Exception {  
                //��ʼִ�к�ʱ����  
            	JacobUtil ja=new JacobUtil();
            	ja.aaa();
                Thread.sleep(1000 * 5);  
                return "�߳�ִ�����.";  
            }  
        };  
          
        try {  
            Future<String> future = exec.submit(call);  
            String obj = future.get(1000 * 3, TimeUnit.MILLISECONDS); //������ʱʱ����Ϊ 1 ��  
            System.out.println("����ɹ�����:" + obj);  
        } catch (TimeoutException ex) {  
            System.out.println("����ʱ��....");  
            ex.printStackTrace();  
        } catch (Exception e) {  
            System.out.println("����ʧ��.");  
            e.printStackTrace();  
        }  
        // �ر��̳߳�  
        exec.shutdown();  
    }
    
}