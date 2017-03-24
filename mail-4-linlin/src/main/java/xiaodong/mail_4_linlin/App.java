package xiaodong.mail_4_linlin;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadFactory;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JScrollBar;
import javax.swing.JScrollPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.border.TitledBorder;

public class App  {
	private static App instance = new App();
	private JFrame frame;
    private JPanel mainPanel;
    private JButton sendMailButton;
    
    //browse excel panel, the excel stores the information of how to send the mail
    private JButton browseExcel;
    private FileChooser excelChooser;
    private JTextField excelFileText;
    private List<ExcelObj> mailReceivers;
    
    //browse attachment panel
    private FileChooser attachChooser;
    private JTextField attachlFileText;
    private JButton attachBrowse;

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
    private App() {
    }

    public static App getInstance() {
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
    	App.getInstance().init();
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
    	
    	//mail attachment input panel
    	attachmentPanel();
    	
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
					excelChooser = new FileChooser("选择表格文件");
					excelFileText.setText(excelChooser.getSelectedFile().getAbsolutePath());
					ExcelReader.RowConverter<ExcelObj> converter = (row)->new ExcelObj(row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10]);
		             ExcelReader<ExcelObj>reader = ExcelReader.builder(ExcelObj.class).converter(converter).withHeader().build();
		             mailReceivers = reader.read(excelChooser.getSelectedFile());
				} catch (Exception e2) {
					e2.printStackTrace();
				}
			}
		});
   }
   
   /**
    * mail attachment input panel
    */
   public void attachmentPanel() {
       JPanel attachPane = new JPanel();
       attachPane.setBorder(new TitledBorder(null, "附件",
               TitledBorder.LEADING, TitledBorder.TOP, null, null));
       attachPane.setBounds(8, 100, 800, 80);
       mainPanel.add(attachPane);
       attachPane.setLayout(new FlowLayout());
       attachlFileText = new JTextField("", 40);
       attachPane.add(attachlFileText);
       attachBrowse = new JButton("附件");
       attachPane.add(attachBrowse);
       attachBrowse.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					attachChooser = new FileChooser("选择附件");
					attachlFileText.setText(attachChooser.getSelectedFile().getAbsolutePath());

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
       sendMailButton = new JButton("发送");
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
					for (ExcelObj obj : mailReceivers) {
						if (!Constant.addAndCheck(obj.getPhone())) {
							String tip = "\r\nalread send mail to " + obj.getName() + "will not send to him again!";
							updateTextArea(tip);
							continue;
						}
						long now = System.currentTimeMillis();
						String beginSend = "sending to " + obj.getName() + "...";
						updateTextArea(beginSend);
						sendStatus = SendStatus.Sending;
						boolean reslult = service.doSendHtmlEmail(obj, attachChooser.getSelectedFile());
						if(!reslult) {
							updateTextArea("fail send to " + obj.getName() + "...");
						}
						sendStatus = SendStatus.DoneOne;
						String end = "\r\nsending to " + obj.getName() + " done!!! 用时:"
								+ (System.currentTimeMillis() - now);
						updateTextArea(end);

					}
					sendStatus = SendStatus.AllDone;
				} catch (Exception e1) {
					logArea.setText("\r\n发送失败， 老婆程序要退出啦");
					System.exit(1);
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
							updateTextArea("\r\n正在发送第"+i+"封邮件。。。");
							Thread.sleep(1000);
							break;
						case DoneOne:
							Thread.sleep(500);
							i++;
							break;
						case AllDone:
							updateTextArea("\r\n全部发送完毕。。。");
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
