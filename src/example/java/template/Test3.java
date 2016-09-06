/**
 * 
 */
package template;

import xgt.util.excel.Config;
import xgt.util.excel.EasyTemplate;
import xgt.util.excel.TemplateFactory;
import xgt.util.excel.model.Pagers;
import xgt.util.excel.model.Record;
import xgt.util.excel.model.RowType;
import xgt.util.excel.templates.SimpleTemplate;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;

/**
 * @author Gavin
 *
 */
public class Test3 {
	public static void main(String[] args) {
		Pagers pagers = new Pagers();
		for (int j = 0; j < 5; j++) {
			Record header = new Test1Model(new Object[]{ "周一", "周二", "周三", "周四", "周五", "周六", "周日" });

			Collection<Record> list = new ArrayList<Record>();
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
			MyPager pager = new MyPager("测试" + j, header, list);
			pagers.addPager(pager);
		}

		pagers.setTitleOnlyFirstPage(false);
		pagers.setLineSpacing(1);

		EasyTemplate t = TemplateFactory.createTemplate(SimpleTemplate.class,pagers);

		Config config = t.getConfig(0);

		config.setDefaultWidth(config.getDefaultWidth()*2);
        config.addRowHeight(config.getDefaultHeight()+10, RowType.TITLE);

		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xls");
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
