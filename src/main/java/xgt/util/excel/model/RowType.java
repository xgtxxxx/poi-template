/**
 * 
 */
package xgt.util.excel.model;

/**
 * @author Gavin
 *
 */
public enum RowType {
	TITLE("title", -1),
	HEADER("header", -2),
	BODY("body", -3),
	FOOTER("footer", -4),
	BLANK("blank", -5),
	DEFAULT("default", -6);
	
	private String name;

	private int num;

	private RowType(String name, int num) {
		this.name = name;
		this.num = num;
	}

	/**
	 * @return the name
	 */
	public String getName() {
		return name;
	}

	public int getNum(){
		return num;
	}
}