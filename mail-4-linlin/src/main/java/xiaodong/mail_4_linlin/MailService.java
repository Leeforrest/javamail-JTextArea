package xiaodong.mail_4_linlin;

import java.io.File;
import java.io.StringWriter;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.internet.MimeUtility;

import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.VelocityEngine;


public class MailService {
    /**
     * Message对象将存储我们实际发送的电子邮件信息，
     * Message对象被作为一个MimeMessage对象来创建并且需要知道应当选择哪一个JavaMail session。
     */
    private MimeMessage message;
    
    /**
     * Session类代表JavaMail中的一个邮件会话。
     * 每一个基于JavaMail的应用程序至少有一个Session（可以有任意多的Session）。
     * 
     * JavaMail需要Properties来创建一个session对象。
     * 寻找"mail.smtp.host"    属性值就是发送邮件的主机
     * 寻找"mail.smtp.auth"    身份验证，目前免费邮件服务器都需要这一项
     */
    private Session session;
    
    /***
     * 邮件是既可以被发送也可以被受到。JavaMail使用了两个不同的类来完成这两个功能：Transport 和 Store。 
     * Transport 是用来发送信息的，而Store用来收信。对于这的教程我们只需要用到Transport对象。
     */
    private Transport transport;
    
    private String mailHost="";
    private String sender_username="";
    private String sender_password="";

    
    private Properties properties = new Properties();
    /*
     * 初始化方法
     */
    public MailService() {
            this.mailHost = "smtp.263.net";
            this.sender_username = "lvlinlin@landsky.cn";
            this.sender_password = "sophia8831";

        properties.setProperty("mail.sender.username", this.sender_username);
        properties.setProperty("mail.sender.password", this.sender_password);
        properties.setProperty("mail.smtp.host", this.mailHost);
        
        session = Session.getInstance(properties);
//        session.setDebug(debug);//开启后有调试信息
        message = new MimeMessage(session);
    }

    /**
     * 发送邮件
     * 
     * @param sendHtml
     *            邮件内容
     * @param receiveUser
     *            收件人地址
     */
	public boolean doSendHtmlEmail(ExcelObj obj, File file) {
		try {
			// 发件人
			// InternetAddress from = new InternetAddress(sender_username);
			// 下面这个是设置发送人的Nick name
			// InternetAddress from = new
			// InternetAddress(MimeUtility.encodeWord("幻影")+"
			// <"+sender_username+">");
			message.setFrom(new InternetAddress(sender_username));
			String content = mailContent(obj);
			// 收件人
			if (obj.getMail() == null) {
				return false;
			}
			InternetAddress to = new InternetAddress(obj.getMail());
			message.setRecipient(Message.RecipientType.TO, to);// 还可以有CC、BCC

			// 邮件主题
			message.setSubject("北京良业环境技术有限公司 面试邀请");

			// 向multipart对象中添加邮件的各个部分内容，包括文本内容和附件
			Multipart multipart = new MimeMultipart();

			// 添加邮件正文
			BodyPart contentPart = new MimeBodyPart();
			contentPart.setContent(content, "text/html;charset=UTF-8");
			multipart.addBodyPart(contentPart);

			// 添加附件的内容
			if (file != null) {
				BodyPart attachmentBodyPart = new MimeBodyPart();
				DataSource source = new FileDataSource(file);
				attachmentBodyPart.setDataHandler(new DataHandler(source));

				// 网上流传的解决文件名乱码的方法，其实用MimeUtility.encodeWord就可以很方便的搞定
				// 这里很重要，通过下面的Base64编码的转换可以保证你的中文附件标题名在发送时不会变成乱码
				// sun.misc.BASE64Encoder enc = new sun.misc.BASE64Encoder();
				// messageBodyPart.setFileName("=?GBK?B?" +
				// enc.encode(attachment.getName().getBytes()) + "?=");

				// MimeUtility.encodeWord可以避免文件名乱码
				attachmentBodyPart.setFileName(MimeUtility.encodeWord(file.getName()));
				multipart.addBodyPart(attachmentBodyPart);
			}

			// 将multipart对象放到message中
			message.setContent(multipart);

			// 保存邮件
			message.saveChanges();

			transport = session.getTransport("smtp");
			// smtp验证，就是你用来发邮件的邮箱用户名密码
			transport.connect(mailHost, sender_username, sender_password);
			// 发送
			transport.sendMessage(message, message.getAllRecipients());
			// System.out.println("send success!");
		} catch (Exception e) {
			e.printStackTrace();
			return false;
		} finally {
			if (transport != null) {
				try {
					transport.close();
				} catch (MessagingException e) {
					e.printStackTrace();
				}
			}
		}
		return true;
	}
    
	private String mailContent(ExcelObj obj) {
		String mv = System.getProperty("user.dir");
		VelocityEngine ve = new VelocityEngine();
//		ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "classpath");
//		ve.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
		properties.setProperty(VelocityEngine.FILE_RESOURCE_LOADER_PATH, mv);;
        
//        properties.setProperty(Velocity.ENCODING_DEFAULT, "GBK");
//        properties.setProperty(Velocity.INPUT_ENCODING, "GBK");
//        properties.setProperty(Velocity.OUTPUT_ENCODING, "GBK");
		ve.init();
		Template t = ve.getTemplate("mail.vm", "UTF-8");
		VelocityContext ctx = new VelocityContext();

		ctx.put("name", obj.getName());
		ctx.put("time", obj.getTime());
		ctx.put("job", obj.getJob());
		StringWriter sw = new StringWriter();

		t.merge(ctx, sw);
		return sw.toString();
	}
    
}