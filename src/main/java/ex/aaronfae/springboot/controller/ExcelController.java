package ex.aaronfae.springboot.controller;

import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.time.DateUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import ex.aaronfae.springboot.excel.entity.Person;
import ex.aaronfae.springboot.excel.util.ExcelUtil;

@Controller
@RequestMapping(value = "/excel")
public class ExcelController {

	@RequestMapping(value = { "", "/" })
	public String index() {
		return "import";
	}

	@RequestMapping("exportExcel")
	public void export(HttpServletResponse response) {

		// 模拟从数据库获取需要导出的数据
		List<Person> personList = new ArrayList<>();
		Person person1 = new Person("路飞", "1", new Date());
		Person person2 = new Person("娜美", "2", DateUtils.addDays(new Date(), 3));
		Person person3 = new Person("索隆", "1", DateUtils.addDays(new Date(), 10));
		Person person4 = new Person("小狸猫", "1", DateUtils.addDays(new Date(), -10));
		personList.add(person1);
		personList.add(person2);
		personList.add(person3);
		personList.add(person4);

		// 导出操作
		ExcelUtil.exportExcel(personList, null, "草帽一伙", Person.class, "海贼王.xls", response);
	}

	@RequestMapping("importExcel")
	@ResponseBody
	public String importExcel(MultipartFile file) {
		// 解析excel
		List<Person> personList = ExcelUtil.importExcel(file, 1, 1, Person.class);
		// String filePath = "D:\\海贼王.xls";
		// 也可以使用String filePath,使用 ExcelUtil.importExcel(String file, Integer
		// titleRows, Integer headerRows, Class<T> pojoClass)导入
		System.out.println("导入数据一共【" + personList.size() + "】行");
		// TODO 保存数据库
		return "导入数据一共【" + personList.size() + "】行";
	}
}
