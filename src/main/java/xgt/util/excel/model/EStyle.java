package xgt.util.excel.model;

import org.apache.poi.ss.usermodel.Font;
import xgt.util.excel.styles.Border;
import xgt.util.excel.utils.StyleDecorate;

public class EStyle{
    public static final String FORMAT_MONEY_DOLLAR = "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)";
    public static final String FORMAT_MONEY_RMB = "_(￥* #,##0.00_);_(￥* (#,##0.00);_(￥* \"-\"??_);_(@_)";
    public static final String FORMAT_PERCENTAGE = "0.00%";
    public static final String FORMAT_DATE = "yyyy-MM-dd";

    public static final short FONT_SIZE_BIGGER = 16;
    public static final short FONT_SIZE_BIG = 14;
    public static final short FONT_SIZE_NORMAL = 12;
    public static final short FONT_SIZE_SMALL = 10;

    public static final short FONT_COLOR_RED = Font.COLOR_RED;

    private Border border;
    private short align;
    private String format;
    private short bgColor;
    private short fontSize = FONT_SIZE_NORMAL;
    private short fontColor;
    private byte fontUnderLine;
    private boolean fontBold;
    private String fontName = "宋体";

    public static EStyle newDefaultTitle(){
        return StyleDecorate.decorateAsTitle(new EStyle());
    }

    public static EStyle newDefaultHeader(){
        return StyleDecorate.decorateAsHeader(new EStyle());
    }

    public Border getBorder() {
        return border;
    }

    public void setBorder(Border border) {
        this.border = border;
    }

    public short getAlign() {
        return align;
    }

    public void setAlign(short align) {
        this.align = align;
    }

    public String getFormat() {
        return format;
    }

    public void setFormat(String format) {
        this.format = format;
    }

    public short getBgColor() {
        return bgColor;
    }

    public void setBgColor(short bgColor) {
        this.bgColor = bgColor;
    }

    public short getFontSize() {
        return fontSize==0?FONT_SIZE_NORMAL:fontSize;
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public short getFontColor() {
        return fontColor;
    }

    public void setFontColor(short fontColor) {
        this.fontColor = fontColor;
    }

    public byte getFontUnderLine() {
        return fontUnderLine;
    }

    public void setFontUnderLine(byte fontUnderLine) {
        this.fontUnderLine = fontUnderLine;
    }

    public boolean isFontBold() {
        return fontBold;
    }

    public void setFontBold(boolean fontBold) {
        this.fontBold = fontBold;
    }

    public String getFontName() {
        return fontName;
    }

    public void setFontName(String fontName) {
        this.fontName = fontName;
    }

    public static EStyle newBorderInstance(){
        final EStyle style = new EStyle();
        style.setBorder(Border.newDefaultInstance());
        return style;
    }
}
