/**
 * 
 */
package template;

import xgt.util.excel.Config;
import xgt.util.excel.EasyTemplate;
import xgt.util.excel.Pager;
import xgt.util.excel.TemplateFactory;
import xgt.util.excel.model.Record;
import xgt.util.excel.templates.SimpleTemplate;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @author Gavin
 *
 */
public class Test2 {
	public static void main(String[] args) {
		String[] header = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };
        Test1Model h = new Test1Model(header);
		List<Test1Model> list = new ArrayList<Test1Model>();
		for (int i = 0; i < 6; i++) {
			Object[] data = new Object[7];
			data[1] = "String" + i;
			data[2] = new Object();
			data[3] = 10 + i;
			data[4] = 20.5 + i;
			data[5] = true;
			data[6] = new Date();
			data[0] = null;
			list.add(new Test1Model(data));
		}

		Pager pager = new MyPager("测试", h, list);

		EasyTemplate t = TemplateFactory.createTemplate(SimpleTemplate.class,pager);

		Config config = t.getConfig(0);
		config.setDefaultWidth(Config.DEFAULT_WIDTH+5);

		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xlsx");
			t.build(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
