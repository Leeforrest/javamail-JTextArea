package xiaodong.mail_4_linlin;

import javax.swing.JTextArea;

public class TextAreaThread implements Runnable{
    //记录日志的文本域
    JTextArea area;
     
    //一条记录
    String newRecord = "";
     
    private static TextAreaThread log;
     
    //单例模式返回日志对象
    public static TextAreaThread getInstance(){
        if(log == null){
            log = new TextAreaThread();
        }
        return log;
    }
     
    //获得新记录 无值则等待
    public synchronized String getNewRecord(){
        try {
            if(newRecord == ""){
                wait();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        String result = newRecord;
        newRecord = "";
         
        notify();
         
        return result;
    }
    //设置新记录 有值则等待
    public synchronized void setNewRecord(String record){
        try {
            if(newRecord != ""){
                wait();
            }else{
                this.newRecord = record;
                notify();
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
     
    @Override
    public void run() {
        while(true){
            String newRecord = getNewRecord();
//        	if("".equals(newRecord)) {
//        		newRecord = "发送中...";//这么做是线程不安全的
//        	}
            //日志框中打印日志
            area = App.logArea;
            area.append(" "+newRecord+"\r\n");
            area.setCaretPosition(area.getText().length());
            try {
				Thread.sleep(1000);
			} catch (InterruptedException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
        }
    }
}
