package xgt.util.excel.templates;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import xgt.util.excel.Reader;
import xgt.util.excel.model.ECell;
import xgt.util.excel.model.ERow;
import xgt.util.excel.model.ESheet;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class SimpleReaderTemplate implements Reader {
    private static final Logger LOGGER = Logger.getLogger(SimpleReaderTemplate.class);

    public static Reader newInstance(){
        return new SimpleReaderTemplate();
    }

    public List<ESheet> read(final InputStream is) throws IOException {
        Workbook workbook = null;
        final List<ESheet> eSheets = new ArrayList<ESheet>();
        try{
            workbook = new XSSFWorkbook(is);
            final List<Sheet> sheets = getSheets(workbook);
            for(final Sheet sheet: sheets){
                eSheets.add(resdSheet(sheet));
            }
        }catch (final IOException e){
            LOGGER.error(e);
            throw e;
        }finally {
            workbook.close();
        }
        return eSheets;
    }
    private ESheet resdSheet(final Sheet sheet) {
        final ESheet eSheet = new ESheet(sheet.getSheetName());
        final int first = sheet.getFirstRowNum();
        final int last = sheet.getLastRowNum();
        for (int i = first; i <= last; i++) {
            final Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            final ERow er = ERow.newReadInstance(i);
            final int colNum = row.getPhysicalNumberOfCells();
            int j = 0;
            int index = 0;
            while (index < colNum) {
                final Cell c = row.getCell(j);
                if(c!=null){
                    final Object value = getCellValue(c);
                    final ECell cell = new ECell(i, j, value);
                    er.addCell(cell);
                    index++;
                }
                j++;
            }
            eSheet.addRow(er);
        }
        return eSheet;
    }

    private List<Sheet> getSheets(final Workbook workbook){
        final List<Sheet> sheets = new ArrayList<Sheet>();
        int index = 0;
        while (true){
            try{
                sheets.add(workbook.getSheetAt(index));
            }catch (final IllegalArgumentException e){
                break;
            }
            index++;
        }
        return sheets;
    }

    private Object getCellValue(final Cell cell) {
        if (cell == null) {
            return "";
        }
        Object strCell = "";
        switch (cell.getCellType()) {
            case HSSFCell.CELL_TYPE_STRING:
                strCell = cell.getStringCellValue();
                break;
            case HSSFCell.CELL_TYPE_BOOLEAN:
                strCell = cell.getBooleanCellValue();
                break;
            case HSSFCell.CELL_TYPE_BLANK:
                break;
            case HSSFCell.CELL_TYPE_NUMERIC:
            case HSSFCell.CELL_TYPE_FORMULA: {
                // 判断当前的cell是否为Date
                if (HSSFDateUtil.isCellDateFormatted(cell)) {
                    strCell = cell.getDateCellValue();
                }else {
                    strCell = cell.getNumericCellValue();
                }
                break;
            }
        }
        return strCell;
    }
}
