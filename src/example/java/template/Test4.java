package template;

import xgt.util.excel.EasyTemplate;
import xgt.util.excel.SheetTemplate;
import xgt.util.excel.model.Record;
import xgt.util.excel.templates.SimpleTemplate;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;

public class Test4 {
    public static void main(final String[] args) throws IllegalAccessException, IOException {
        final EasyTemplate template = new SimpleTemplate();
        for(int j=0; j<3; j++){
            final SheetTemplate sheet = new SheetTemplate();
            sheet.setName("name"+j);
            final Record header = new Test1Model(new Object[]{ "周一", "周二", "周三", "周四", "周五", "周六", "周日" });
            final Collection<Record> list = new ArrayList<Record>();
            for (int i = 0; i < 6; i++) {
                final Object[] data = new Object[7];
                data[1] = "String" + i;
                data[2] = new Object();
                data[3] = 10 + i;
                data[4] = 20.5 + i;
                data[5] = true;
                data[6] = new Date();
                data[0] = null;
                list.add(new Test1Model(data));
            }
            sheet.addData(header, list);
            template.addSheet(sheet);
        }

        final FileOutputStream fos = new FileOutputStream("D:/text.xls");
        template.build(fos);
    }
}
