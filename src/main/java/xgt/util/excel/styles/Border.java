/**
 * 
 */
package xgt.util.excel.styles;

import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author Gavin
 *
 */
public class Border {
	private Map<Directive,BorderStyle> styles = new HashMap<Directive,BorderStyle>();
	
	public void addStyle(Directive dir,short type,short color){
		this.styles.put(dir, new BorderStyle(dir,type,color));
	}
	
	public void addStyle(Directive dir, short type){
		this.addStyle(dir,type,(short)0);
	}
	
	public BorderStyle getBorderStyle(Directive dir){
		return styles.get(dir);
	}
	
	public short getColor(Directive dir){
		return this.styles.get(dir).getColor();
	}
	
	public short getType(Directive dir){
		return this.styles.get(dir).getType();
	}
	
	public static Border newDefaultInstance(){
		Border border = new Border();
		border.addStyle(Directive.TOP, CellStyle.BORDER_THIN);
		border.addStyle(Directive.LEFT, CellStyle.BORDER_THIN);
		border.addStyle(Directive.RIGHT, CellStyle.BORDER_THIN);
		border.addStyle(Directive.BOTTOM, CellStyle.BORDER_THIN);
		return border;
	}
}
