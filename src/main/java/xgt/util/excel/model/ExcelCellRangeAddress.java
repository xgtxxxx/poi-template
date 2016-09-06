package xgt.util.excel.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * 合并的单元格地址
 * @author xgt
 *
 */
public class ExcelCellRangeAddress extends CellRangeAddress {
	//所有合并单元格集合
	private static List<ExcelCellRangeAddress> ranges = new ArrayList<ExcelCellRangeAddress>();
	//合并单元格样式
	private CellStyle cellStyle;

	public ExcelCellRangeAddress(int firstRow, int lastRow, int firstCol,
								 int lastCol) {
		super(firstRow, lastRow, firstCol, lastCol);
	}
	/**
	 * 获取合并单元格的样式
	 * @return
	 */
	public CellStyle getCellStyle() {
		return cellStyle;
	}
	/**
	 * 设置合并单元格的样式
	 * @param cellStyle
	 */
	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}
	
	public static void addRanges(ExcelCellRangeAddress ca) {
		ranges.add(ca);
	}
	
	public static List<ExcelCellRangeAddress> getRanges(){
		return ranges;
	}
	
	public static void clear(){
		ranges.clear();
	}
}
