/**
 * 
 */
package xgt.util.excel;

import java.util.ArrayList;
import java.util.List;

/**
 * @author Gavin
 *
 */
public final class Region {
	/**
	 * @param startRow
	 * @param endRow
	 * @param startColumn
	 * @param endColumn
	 */
	public Region(int startRow, int endRow, int startColumn, int endColumn) {
		super();
		this.startRow = startRow;
		this.endRow = endRow;
		this.startColumn = startColumn;
		this.endColumn = endColumn;
	}
	private int startRow;
	private int endRow;
	private int startColumn;
	private int endColumn;
	
	public List<String> getRegionDetails(){
		List<String> list = new ArrayList<String>();
		for(int i=startRow; i<=endRow; i++){
			for(int j=startColumn; j<=endColumn; j++){
				String s = i+"-"+j;
				list.add(s);
			}
		}
		return list;
	}
	
	/**
	 * @return the startRow
	 */
	public int getStartRow() {
		return startRow;
	}
	/**
	 * @param startRow the startRow to set
	 */
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}
	/**
	 * @return the endRow
	 */
	public int getEndRow() {
		return endRow;
	}
	/**
	 * @param endRow the endRow to set
	 */
	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}
	/**
	 * @return the startColumn
	 */
	public int getStartColumn() {
		return startColumn;
	}
	/**
	 * @param startColumn the startColumn to set
	 */
	public void setStartColumn(int startColumn) {
		this.startColumn = startColumn;
	}
	/**
	 * @return the endColumn
	 */
	public int getEndColumn() {
		return endColumn;
	}
	/**
	 * @param endColumn the endColumn to set
	 */
	public void setEndColumn(int endColumn) {
		this.endColumn = endColumn;
	}

}
