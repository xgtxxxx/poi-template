package xgt.util.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import xgt.util.excel.model.EStyle;

public interface StyleAdaptor {
    public CellStyle adaptTo(final Workbook wb, final EStyle style);
}
