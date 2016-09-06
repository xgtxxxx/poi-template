/**
 *
 */
package xgt.util.excel.utils;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.IndexedColors;
import xgt.util.excel.model.EStyle;
import xgt.util.excel.styles.Border;

/**
 * @author Gavin
 *
 */
public class StyleDecorate {

    private StyleDecorate(){};

    public static EStyle addBorder(EStyle style){
        style.setBorder(Border.newDefaultInstance());
        return style;
    }

    public static EStyle addBorder(EStyle style,Border border){
        style.setBorder(border);
        return style;
    }

    public static EStyle decorateLeft(EStyle style){
        return decorateAlign(style, CellStyle.ALIGN_LEFT);
    }

    public static EStyle decorateCenter(EStyle style){
        return decorateAlign(style,CellStyle.ALIGN_CENTER);
    }

    public static EStyle decorateRight(EStyle style){
        return decorateAlign(style,CellStyle.ALIGN_RIGHT);
    }

    public static EStyle decorateAlign(EStyle style,short align){
        style.setAlign(align);
        return style;
    }

    public static EStyle decorateJustify(EStyle style){
        style.setAlign(CellStyle.ALIGN_JUSTIFY);
        return style;
    }

    public static EStyle decorateAsTitle(EStyle style){
        FontDecorate.decorateDefaultTitle(style);
        style.setAlign(CellStyle.ALIGN_CENTER);
        return style;
    }

    public static EStyle decorateAsHeader(EStyle style){
        FontDecorate.decorateBold(style);
        style.setAlign(CellStyle.ALIGN_CENTER);
        style.setBorder(Border.newDefaultInstance());
        return style;
    }

    public static EStyle decorateAsFooter(EStyle style){
        FontDecorate.decorateBold(style);
        style.setAlign(CellStyle.ALIGN_CENTER);
        return style;
    }

    public static EStyle decorateBgColor(EStyle style, short color){
        style.setBgColor(color);
        return style;
    }

    public static EStyle decorateBgYellow(EStyle style){
        return decorateBgColor(style,IndexedColors.YELLOW.getIndex());
    }

    public static EStyle decorateBgGrey(EStyle style){
        return decorateBgColor(style,IndexedColors.GREY_25_PERCENT.getIndex());
    }

    public static EStyle decorateBgBlue(EStyle style){
        return decorateBgColor(style,IndexedColors.BLUE.getIndex());
    }

}