package merge4chong;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import xiaodong.mail_4_linlin.*;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;

/**
 * 有一个文件夹，文件夹下有若干个包含excel的文件夹和一个excel文件，要将所有文件夹下的excel合并到外边这个excel中
 */
public class MergeApplication  {
    private static MergeApplication instance = new MergeApplication();
    private JFrame frame;
    private JPanel mainPanel;
    private JButton sendMailButton;

    //browse excel
    private JButton browseExcel;
    private FileChooser excelChooser;
    private JTextField excelFileText;
    private List<File> excels = new ArrayList<>();
    private File destExcel;


    //log area
    public static JTextArea logArea;
    public JScrollPane logScroll;

    private SendStatus sendStatus;//单封邮件发送成功/失败
    private ExecutorService service = Executors.newCachedThreadPool(new ThreadFactory() {

        @Override
        public Thread newThread(Runnable r) {
            return new Thread(r, "output");
        }
    });
    private MergeApplication() {

    }

    public static MergeApplication getInstance() {
        return instance;
    }

    enum SendStatus{
        Begin,
        Sending,
        DoneOne,
        AllDone,
    }

    public static void main(String[]args) {
        try {
            UIManager.setLookAndFeel("com.jtattoo.plaf.smart.SmartLookAndFeel");
        } catch(Exception e) {
            e.printStackTrace();
        }
        MergeApplication.getInstance().init();
    }


    /**
     * Create the frame.
     */
    public void init() {

        //主界面
        mainFrame();

        //mainPanel
        mainPanel();

        //excel input panel
        excelInput();

        //should create a new thread to execute the button event, otherwise the main frame will be suspended,
        //then the JTextArea will not show logs real-timely
        sendMailButton();

        //add JTextArea to show logs
        logPanel();

        frame.setVisible(true);
    }

    /**
     * 初始化主机面
     */
    public void mainFrame() {
        frame  = new JFrame("工具");
        int appWidth = 828;
        int appHeight = 800;
        int windowWidth =  frame.getWidth(); // 获得窗口宽
        int windowHeight = frame.getHeight(); // 获得窗口高
        Toolkit kit = Toolkit.getDefaultToolkit(); // 定义工具包
        Dimension screenSize = kit.getScreenSize(); // 获取屏幕的尺寸
        int screenWidth = screenSize.width; // 获取屏幕的宽
        int screenHeight = screenSize.height; // 获取屏幕的高
        frame.setLocation(screenWidth / 2 - windowWidth / 2 -appWidth/2, screenHeight / 2 - windowHeight / 2 - appHeight/2);// 设置窗口居中显示
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(appWidth, appHeight);
    }

    /**
     * 初始化主panel
     */
    public void mainPanel() {
        mainPanel = new JPanel();
        mainPanel.setBorder(new TitledBorder(UIManager
                .getBorder("TitledBorder.border"), "工具",
                TitledBorder.LEADING, TitledBorder.TOP, null,
                new Color(0, 0, 0)));
        frame.setContentPane(mainPanel);
        mainPanel.setLayout(null);
    }

    /**
     * excel input panel
     */
    public void excelInput() {
        JPanel panel = new JPanel();
        panel.setBorder(new TitledBorder(null, "excel",
                TitledBorder.LEADING, TitledBorder.TOP, null, null));
        panel.setBounds(8, 20, 800, 80);
        mainPanel.add(panel);
        panel.setLayout(new FlowLayout());
        excelFileText = new JTextField("", 40);
        panel.add(excelFileText);
        browseExcel = new JButton("excel");
        panel.add(browseExcel);
        browseExcel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                try {
                    excelChooser = new FileChooser("选择表格最终要合并到的excel");
                    destExcel = excelChooser.getSelectedFile();
                    excelFileText.setText(destExcel.getAbsolutePath());
                    File[]files = excelChooser.getSelectedFile().getParentFile().listFiles();

                    updateTextArea("待合并表格:\n");
                    for (File file : files) {
                        if(!file.isDirectory()) {
                            continue;
                        }
                        File[] listfiles = file.listFiles();
                        for (File excel : listfiles) {
                            if(excel.getAbsolutePath().endsWith(".xlsx") || excel.getAbsolutePath().endsWith(".xlsm")) {
                                excels.add(excel);
                                updateTextArea(excel.getAbsolutePath()+"\n");
                            }
                        }
                    }
                    updateTextArea("最终表格:\n"+excelChooser.getSelectedFile().getParentFile().getAbsolutePath() + File.separator + "test" +"\n");
                } catch (Exception e2) {
                    e2.printStackTrace();
                }
            }
        });
    }


    /**
     * add send mail button
     */
    public void sendMailButton() {
        JPanel sendPanel = new JPanel();
        sendPanel.setBounds(8, 198, 800, 30);
        mainPanel.add(sendPanel);
        sendMailButton = new JButton("合并");
        sendPanel.add(sendMailButton);
        sendMailButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                sendExe();
            }
        });
    }

    /**
     * add JTextArea to show logs
     */
    public void logPanel() {
        JPanel show = new JPanel();
        show.setBorder(new TitledBorder(null, "显示",
                TitledBorder.LEADING, TitledBorder.TOP, null, null));
        show.setBounds(8, 240, 800, 400);
        mainPanel.add(show);
        show.setLayout(new BorderLayout());

        logArea = new JTextArea("");
        logScroll = new JScrollPane(logArea); // 滚动面板.
        logArea.setAutoscrolls(true);
        logArea.setLineWrap(true);
        logArea.setWrapStyleWord(true);
        logScroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);

        show.add(logScroll, BorderLayout.CENTER);
    }


    private void updateTextArea(final String text) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                logArea.append(text);
                logArea.paintImmediately(logArea.getBounds());
                JScrollBar   sbar= logScroll.getVerticalScrollBar();
                sbar.setValue(sbar.getMaximum());

            }
        });
    }

    private void sendExe() {
        service.submit(new Runnable() {

            @Override
            public void run() {
                MailService service = new MailService();
                try {
                    sendStatus = SendStatus.Begin;
                    MergeExcel.mergeXSSFWorkbooks(logArea, destExcel, excels);
                } catch (Exception e1) {

                    logArea.setText("\r\n失败， 程序要退出啦\r\n" + e1);
//                    System.exit(1);
                }

            }
        });
        service.submit(new Runnable() {

            @Override
            public void run() {
                try {
                    int i=1;
                    sendStatus = SendStatus.Begin;
                    while (true) {
                        switch (sendStatus) {
                            case Begin:
                                Thread.sleep(500);
                                break;
                            case Sending:
                                updateTextArea("\r\n正在合并第"+i+"个文件。。。");
                                Thread.sleep(1000);
                                break;
                            case DoneOne:
                                Thread.sleep(500);
                                i++;
                                break;
                            case AllDone:
                                updateTextArea("\r\n全部完毕。。。");
                                return;
                            default:
                                break;
                        }
                    }
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

}
