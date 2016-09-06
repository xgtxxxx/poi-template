/**
 * 
 */
package xgt.util.excel.model;

/**
 * @author Gavin
 *
 */
public final class ECell {
	
	private int rowIndex;
	
	private int columnIndex;
	
	private Object value;

	/**
	 * @param rowIndex2
	 * @param columnIndex2
	 * @param value2
	 */
	public ECell(int rowIndex2, int columnIndex2, Object value2) {
		this.rowIndex = rowIndex2;
		this.columnIndex = columnIndex2;
		this.value = value2;
	}
	
	public ECell() {
	}

	/**
	 * @return the rowIndex
	 */
	public int getRowIndex() {
		return rowIndex;
	}

	/**
	 * @param rowIndex the rowIndex to set
	 */
	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	/**
	 * @return the columnIndex
	 */
	public int getColumnIndex() {
		return columnIndex;
	}

	/**
	 * @param columnIndex the columnIndex to set
	 */
	public void setColumnIndex(int columnIndex) {
		this.columnIndex = columnIndex;
	}

	/**
	 * @return the value
	 */
	public Object getValue() {
		return value;
	}

	/**
	 * @param value the value to set
	 */
	public void setValue(Object value) {
		this.value = value;
	}

	@Override
	public String toString() {
		return "ECell{" +
				"rowIndex=" + rowIndex +
				", columnIndex=" + columnIndex +
				", value=" + value +
				'}';
	}
}
