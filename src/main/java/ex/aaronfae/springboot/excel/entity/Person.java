package ex.aaronfae.springboot.excel.entity;

import java.util.Date;

import cn.afterturn.easypoi.excel.annotation.Excel;

public class Person {

	@Excel(name = "姓名", orderNum = "0")
	private String name;

	@Excel(name = "性别", replace = { "男_1", "女_2" }, orderNum = "1")
	private String sex;

	@Excel(name = "生日", format = "yyyy-MM-dd", orderNum = "2")
	private Date birthday;

	public Person() {
	}

	public Person(String name, String sex, Date birthday) {
		this.name = name;
		this.sex = sex;
		this.birthday = birthday;
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

	public Date getBirthday() {
		return birthday;
	}

	public void setBirthday(Date birthday) {
		this.birthday = birthday;
	}
}
