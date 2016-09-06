package template;

import xgt.util.excel.Reader;
import xgt.util.excel.model.ECell;
import xgt.util.excel.model.ERow;
import xgt.util.excel.model.ESheet;
import xgt.util.excel.templates.SimpleReaderTemplate;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class ReadTest {
    public static void main(final String[] args) throws IOException {
        final Reader reader = SimpleReaderTemplate.newInstance();
        final InputStream is = new FileInputStream(new File("D:/text.xls"));
        final List<ESheet> sheets = reader.read(is);
        for(final ESheet sheet: sheets){
            for(final ERow row: sheet.getRows()){
                for(final ECell cell: row.getCells()){
                    System.out.print(cell);
                    System.out.println("----"+cell.getValue().getClass());
                }
            }
        }
    }
}
