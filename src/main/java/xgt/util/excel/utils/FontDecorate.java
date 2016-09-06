/**
 * 
 */
package xgt.util.excel.utils;

import org.apache.poi.ss.usermodel.Font;
import xgt.util.excel.model.EStyle;

/**
 * @author Gavin
 *
 */
public class FontDecorate {
	
	public static final short FONT_SIZE_BIGGER = 16;

	public static final short FONT_SIZE_BIG = 14;

	public static final short FONT_SIZE_NORMAL = 12;

	public static final short FONT_SIZE_SMALL = 10;
	
	public static EStyle decorateColor(EStyle style, short color){
		style.setFontColor(color);
		return style;
	}
	
	public static EStyle decorateRed(EStyle font){
		return decorateColor(font, Font.COLOR_RED);
	}
	
	public static EStyle decorateUnderline(EStyle font, byte underline){
		font.setFontUnderLine(underline);
		return font;
	}
	
	public static EStyle decorateDefaultUnderline(EStyle font){
		return decorateUnderline(font,Font.U_SINGLE);
	}
	
	public static EStyle decorateDefaultFamily(EStyle font){
		font.setFontName("宋体");
		return font;
	}
	
	public static EStyle decorateSize(EStyle font,short size){
		font.setFontSize(size);
		return font;
	}
	
	public static EStyle decorateBigSize(EStyle font){
		return decorateSize(font, FONT_SIZE_BIG);
	}
	
	public static EStyle decorateBold(EStyle font){
		font.setFontBold(true);
		return font;
	}
	
	public static EStyle decorateRedBold(EStyle font){
		decorateBold(font);
		decorateRed(font);
		return font;
	}
	
	public static EStyle decorateRedBoldUnderline(EStyle font){
		decorateBold(font);
		decorateRed(font);
		decorateDefaultUnderline(font);
		return font;
	}
	
	public static EStyle decorateRedUnderline(EStyle font){
		decorateRed(font);
		decorateDefaultUnderline(font);
		return font;
	}
	
	public static EStyle decorateBoldUnderline(EStyle font){
		decorateBold(font);
		decorateDefaultUnderline(font);
		return font;
	}
	
	public static EStyle decorateDefaultTitle(EStyle font){
		decorateBold(font);
		decorateBigSize(font);
		return font;
	}
}	
