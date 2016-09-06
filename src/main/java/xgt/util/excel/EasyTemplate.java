package xgt.util.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.omg.CosNaming.NamingContextPackage.NotFound;
import xgt.util.excel.model.ECell;
import xgt.util.excel.model.ERow;
import xgt.util.excel.model.RowType;

import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public abstract class EasyTemplate {
    private List<SheetTemplate> sheets = new ArrayList<SheetTemplate>();

    protected static final String DEFAULT_CELL_STYLE_KEY = "-";

    public List<SheetTemplate> getSheets() {
        return sheets;
    }

    public void addSheet(final SheetTemplate sheetTemplate){
        this.sheets.add(sheetTemplate);
    }

    public void setSheets(final List<SheetTemplate> sheets) {
        this.sheets = sheets;
    }

    public Config getConfig(final int sheetIndex){
        return this.sheets.get(sheetIndex).getSheetStyle();
    }

    public abstract void build(OutputStream os) throws IOException;

    public abstract void initStyles();

    protected abstract CellStyle getCellStyle(final int sheet, final ERow er, final ECell ec);

    protected abstract Workbook getWorkbook();
}
