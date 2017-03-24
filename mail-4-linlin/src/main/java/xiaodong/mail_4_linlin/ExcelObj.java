package xiaodong.mail_4_linlin;

/**
 * Created by Forrest on 2017/3/16.
 */
public class ExcelObj {
    //序号	岗位	姓名	性别	年龄	面试时间	是否在职	联系电话	邮箱	负责人	剩余简历数
    String id;
    String job;
    String name;
    String sex;
    String age;
    String time;
    String onDuty;
    String phone;
    String mail;
    String onCharge;
    String resume;
    String sendSuccess = "发送失败";


    public ExcelObj(String id, String job, String name, String sex, String age, String time, String onDuty,
                    String phone, String mail, String onCharge, String resume) {
        super();
        this.id = id;
        this.job = job;
        this.name = name;
        this.sex = sex;
        this.age = age;
        this.time = time;
        this.onDuty = onDuty;
        this.phone = phone;
        this.mail = mail;
        this.onCharge = onCharge;
        this.resume = resume;
    }
    
    public String getId() {
        return id;
    }
    public void setId(String id) {
        this.id = id;
    }
    public String getJob() {
        return job;
    }
    public void setJob(String job) {
        this.job = job;
    }
    public String getName() {
        return name;
    }
    public void setName(String name) {
        this.name = name;
    }
    public String getSex() {
        return sex;
    }
    public void setSex(String sex) {
        this.sex = sex;
    }
    public String getAge() {
        return age;
    }
    public void setAge(String age) {
        this.age = age;
    }
    public String getTime() {
        return time;
    }
    public void setTime(String time) {
        this.time = time;
    }
    public String getOnDuty() {
        return onDuty;
    }
    public void setOnDuty(String onDuty) {
        this.onDuty = onDuty;
    }
    public String getPhone() {
        return phone;
    }
    public void setPhone(String phone) {
        this.phone = phone;
    }
    public String getMail() {
        return mail;
    }
    public void setMail(String mail) {
        this.mail = mail;
    }
    public String getOnCharge() {
        return onCharge;
    }
    public void setOnCharge(String onCharge) {
        this.onCharge = onCharge;
    }
    public String getResume() {
        return resume;
    }
    public void setResume(String resume) {
        this.resume = resume;
    }
    public String getSendSuccess() {
		return sendSuccess;
	}

	public void setSendSuccess(String sendSuccess) {
		this.sendSuccess = sendSuccess;
	}

	@Override
    public String toString() {
        return  "\r\n"
        		+ "姓名 :------>" + name + "\r\n" +
        		   "职位 :------>" + job + "\r\n" +
        		   "时间 :------>" + time + "\r\n";
    }

}
