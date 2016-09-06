/**
 * 
 */
package xgt.util.excel.model;

import java.util.Collection;
import java.util.HashMap;
import java.util.Map;

/**
 * @author Gavin
 *
 */
public final class ERow {
	
	private int rowIndex;

	private RowType type = RowType.DEFAULT;

	private Map<Integer, ECell> cells = new HashMap<Integer, ECell>();

	public ECell getCell(int columnIndex) {
		return cells.get(columnIndex);
	}

	public Collection<ECell> getCells() {
		return cells.values();
	}

	public void addCell(ECell cell) {
		this.cells.put(cell.getColumnIndex(), cell);
	}

	private ECell createCell(Object value, int columnIndex) {
		return new ECell(this.rowIndex, columnIndex, value);
	}

	public void addCell(Object value, int columnIndex) {
		addCell(createCell(value, columnIndex));
	}

	/**
	 * @return the rowIndex
	 */
	public int getRowIndex() {
		return rowIndex;
	}

	/**
	 * @param rowIndex
	 *            the rowIndex to set
	 */
	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}
	
	public static ERow newBlankInstance(int rowIndex) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		row.setType(RowType.BLANK);
		return row;
	}

	public static ERow newReadInstance(int rowIndex) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		return row;
	}
	
	public static ERow newTitleInstance(String title, int rowIndex) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		row.addCell(title, 0);
		row.setType(RowType.TITLE);
		return row;
	}
	
	public static ERow newHeaderInstance(Object[] header,int rowIndex) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		row.setType(RowType.HEADER);
		int index = 0;
		for (Object value : header) {
			row.addCell(value, index++);
		}
		return row;
	}

	public static ERow newBodyInstance(Object[] datas, int rowIndex) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		row.setType(RowType.BODY);
		int index = 0;
		for (Object value : datas) {
			row.addCell(value, index++);
		}
		return row;
	}

	public static ERow newInstance(Object[] datas, int rowIndex, RowType type) {
		ERow row = new ERow();
		row.setRowIndex(rowIndex);
		int index = 0;
		for (Object value : datas) {
			row.addCell(value, index++);
		}
		row.setType(type);
		return row;
	}

	/**
	 * @return the type
	 */
	public RowType getType() {
		return type;
	}

	/**
	 * @param type
	 *            the type to set
	 */
	public void setType(RowType type) {
		this.type = type;
	}

}
