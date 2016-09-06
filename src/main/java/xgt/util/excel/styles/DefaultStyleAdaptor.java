package xgt.util.excel.styles;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import xgt.util.excel.StyleAdaptor;
import xgt.util.excel.model.EStyle;

public class DefaultStyleAdaptor implements StyleAdaptor {
    public CellStyle adaptTo(final Workbook wb, final EStyle style) {
        final CellStyle cellStyle = wb.createCellStyle();
        if(style.getBorder()!=null){
            cellStyle.setBorderBottom(style.getBorder().getBorderStyle(Directive.BOTTOM).getType());
            cellStyle.setBorderTop(style.getBorder().getBorderStyle(Directive.TOP).getType());
            cellStyle.setBorderLeft(style.getBorder().getBorderStyle(Directive.LEFT).getType());
            cellStyle.setBorderRight(style.getBorder().getBorderStyle(Directive.RIGHT).getType());
            cellStyle.setBottomBorderColor(style.getBorder().getBorderStyle(Directive.BOTTOM).getColor());
            cellStyle.setTopBorderColor(style.getBorder().getBorderStyle(Directive.TOP).getColor());
            cellStyle.setRightBorderColor(style.getBorder().getBorderStyle(Directive.RIGHT).getColor());
            cellStyle.setLeftBorderColor(style.getBorder().getBorderStyle(Directive.LEFT).getColor());
        }
        if(style.getAlign()!=0){
            cellStyle.setAlignment(style.getAlign());
        }
        if(StringUtils.isNotEmpty(style.getFormat())){
            cellStyle.setDataFormat(wb.createDataFormat().getFormat(style.getFormat()));
        }
        if(style.getBgColor()!=0){
            cellStyle.setFillForegroundColor(style.getBgColor());
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        }
        final Font font = wb.createFont();
        font.setBold(style.isFontBold());
        font.setFontHeightInPoints(style.getFontSize());
        font.setColor(style.getFontColor());
        if(style.getFontUnderLine()!=0){
            font.setUnderline(style.getFontUnderLine());
        }
        font.setFontName(style.getFontName());
        cellStyle.setFont(font);
        return cellStyle;
    }
}
