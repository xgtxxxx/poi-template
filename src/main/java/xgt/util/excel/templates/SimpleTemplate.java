package xgt.util.excel.templates;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import xgt.util.excel.Config;
import xgt.util.excel.EasyTemplate;
import xgt.util.excel.Region;
import xgt.util.excel.SheetTemplate;
import xgt.util.excel.StyleAdaptor;
import xgt.util.excel.model.ECell;
import xgt.util.excel.model.ERow;
import xgt.util.excel.model.ESheet;
import xgt.util.excel.model.EStyle;
import xgt.util.excel.model.ExcelCellRangeAddress;
import xgt.util.excel.model.RowType;
import xgt.util.excel.styles.Border;
import xgt.util.excel.styles.DefaultStyleAdaptor;
import xgt.util.excel.utils.StyleDecorate;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class SimpleTemplate extends EasyTemplate {
    private static final int ROW_TITLE = -1;

    private Workbook wb;

    private final Map<String, CellStyle> cellStyleMap = new HashMap<String, CellStyle>();

    @Override
    public void build(final OutputStream os) throws IOException {
        try{
            wb = getWorkbook();
            initStyles();
            for(int i=0; i<getSheets().size(); i++){
                final SheetTemplate template = getSheets().get(i);
                final ESheet data = template.getSheetData();
                final Config config = template.getSheetStyle();
                final Sheet sheet = wb.createSheet(data.getName()==null?"sheet"+(i+1):data.getName());
                initSheetWidthConfig(sheet, config);
                buildRows(sheet, data, config, i);
                setMerge(sheet, template.getRegions());
            }
            wb.write(os);
        }finally {
            if(wb!=null){
                wb.close();
            }
        }
    }

    private void initDefault(final int sheet, final StyleAdaptor adaptor){
        cellStyleMap.put(getKey(sheet, RowType.TITLE.getName()), adaptor.adaptTo(wb, StyleDecorate.decorateAsTitle(new EStyle())));
        cellStyleMap.put(getKey(sheet, RowType.HEADER.getName()), adaptor.adaptTo(wb, StyleDecorate.decorateAsHeader(new EStyle())));
        final EStyle style = new EStyle();
        style.setBorder(Border.newDefaultInstance());
        style.setFontSize(EStyle.FONT_SIZE_NORMAL);
        cellStyleMap.put(getKey(sheet, RowType.BODY.getName()), adaptor.adaptTo(wb, style));
    }

    @Override
    public void initStyles() {
        final StyleAdaptor adaptor = new DefaultStyleAdaptor();
        for(int i=0; i<getSheets().size(); i++){
            final SheetTemplate template = getSheets().get(i);
            final Config config = template.getSheetStyle();
            initDefault(i, adaptor);
            for(final RowType key: config.rowTypeStyleKeys()){
                cellStyleMap.put(i+"-"+key.getName(), adaptor.adaptTo(wb, config.getRowTypeStyle(key)));
            }
            for(final Region key: config.regionStyleKeys()){
                final CellStyle style = adaptor.adaptTo(wb, config.getRegionStyle(key));
                for(int row=key.getStartRow(); row<=key.getEndRow(); row++){
                    for(int column=key.getStartColumn(); column<=key.getEndColumn(); column++){
                        cellStyleMap.put(getKey(i, row, column), style);
                    }
                }
            }
        }
        cellStyleMap.put(DEFAULT_CELL_STYLE_KEY, wb.createCellStyle());
    }

    private String getKey(final int sheetIndex, final int rowIndex, final int columnIndex){
        final StringBuffer sb = new StringBuffer();
        sb.append(sheetIndex).append('-').append(rowIndex).append('-').append(columnIndex);
        return sb.toString();
    }

    private String getKey(final int sheetIndex, final String type){
        final StringBuffer sb = new StringBuffer();
        sb.append(sheetIndex).append('-').append(type);
        return sb.toString();
    }

    @Override
    public CellStyle getCellStyle(final int sheet, final ERow er, final ECell ec) {
        CellStyle style = cellStyleMap.get(getKey(sheet, er.getRowIndex(), ec.getColumnIndex()));
        if(style==null){
            style = cellStyleMap.get(getKey(sheet, er.getType().getName()));
        }
        return style==null?cellStyleMap.get(DEFAULT_CELL_STYLE_KEY):style;
    }

    @Override
    protected Workbook getWorkbook() {
        if(this.wb==null){
            this.wb = new SXSSFWorkbook(500);
        }
        return wb;
    }

    private void setMerge(final Sheet sheet, final List<Region> regions){
        if(regions==null){
            return;
        }
        for (final Region region : regions) {
            final CellRangeAddress cra = new ExcelCellRangeAddress(region.getStartRow(),region.getEndRow(),region.getStartColumn(),region.getEndColumn());
            sheet.addMergedRegion(cra);
        }
    }

    private void initSheetWidthConfig(final Sheet sheet, final Config config){
        //此处会导致excel报：有不可读取的内容。
//		sheet.setDefaultColumnWidth(config.getDefaultWidth()*256);
        sheet.setDefaultRowHeightInPoints(config.getDefaultHeight());
        for (final int index : config.getKeysOfWidth()) {
            sheet.setColumnWidth(index, config.getColumnWidth(index)*256);
        }
    }

    private void buildRows(final Sheet sheet, final ESheet data, final Config config, final int sheetIndex){
        for (final ERow er : data.getRows()) {
            final Row row = sheet.createRow(er.getRowIndex());
            Float rowHeight = config.getRowHeight(er.getRowIndex());
            if(rowHeight==null){
                rowHeight = config.getRowHeight(er.getType());
            }
            row.setHeightInPoints(rowHeight);
            buildCells(row,er,sheetIndex);
        }
    }

    private void buildCells(final Row row, final ERow er, final int sheetIndex){
        for (final ECell ec : er.getCells()) {
            final CellStyle style = getCellStyle(sheetIndex, er, ec);
            final Cell cell = row.createCell(ec.getColumnIndex());
            final Object v = ec.getValue();
            if(v instanceof Boolean){
                cell.setCellValue((Boolean)v);
            }else if(v instanceof Number){
                cell.setCellValue(((Number)v).doubleValue());
            }else if(v instanceof Date){
                cell.setCellValue((Date)v);
            }else if(v instanceof String){
                cell.setCellValue((String)v);
            }else if(v instanceof RichTextString){
                cell.setCellValue((RichTextString)v);
            }else{
                cell.setCellValue(String.valueOf(v));
            }
            cell.setCellStyle(style);
        }
    }
}
