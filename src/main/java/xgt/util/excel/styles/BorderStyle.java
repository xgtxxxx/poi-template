/**
 * 
 */
package xgt.util.excel.styles;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * @author Gavin
 *
 */
public class BorderStyle {
	
	private Directive dir;
	
	private short type = CellStyle.BORDER_THIN;
	
	private short color;
	
	/**
	 * @param dir
	 * @param type
	 * @param color
	 */
	public BorderStyle(Directive dir, short type, short color) {
		super();
		this.dir = dir;
		this.type = type;
		this.color = color;
	}

	/**
	 * @return the dir
	 */
	public Directive getDir() {
		return dir;
	}

	/**
	 * @param dir the dir to set
	 */
	public void setDir(Directive dir) {
		this.dir = dir;
	}

	/**
	 * @return the type
	 */
	public short getType() {
		return type;
	}

	/**
	 * @param type the type to set
	 */
	public void setType(short type) {
		this.type = type;
	}

	/**
	 * @return the color
	 */
	public short getColor() {
		return color;
	}

	/**
	 * @param color the color to set
	 */
	public void setColor(short color) {
		this.color = color;
	}
	
}
