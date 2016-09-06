/**
 * 
 */
package xgt.util.excel;

import java.util.Collection;
import java.util.EnumMap;
import java.util.HashMap;
import java.util.Map;

import xgt.util.excel.model.EStyle;
import xgt.util.excel.model.RowType;

/**
 * @author Gavin
 *
 */
public class Config {
	public static final int DEFAULT_WIDTH = 8;
	
	public static final float DEFAULT_HEIGHT = 12.75f;

	private final Map<Integer,Integer> widths = new HashMap<Integer,Integer>();
	
	private final Map<Integer,Float> heigths= new HashMap<Integer,Float>();
	
	private final Map<Region, EStyle> regionStyles = new HashMap<Region, EStyle>();

	private final Map<RowType, EStyle> rowTypeStyles = new EnumMap<RowType, EStyle>(RowType.class);

	private int defaultWidth = DEFAULT_WIDTH;
	
	private float defaultHeight = DEFAULT_HEIGHT;
	
	public int getColumnWidth(final int index){
		final Integer w = widths.get(index);
		return w==null?defaultWidth:w;
	}
	
	public Float getRowHeight(final int index){
		return this.heigths.get(index);
	}

	public Float getRowHeight(final RowType type){
		return this.heigths.get(type.getNum())==null?getDefaultHeight():this.heigths.get(type.getNum());
	}
	
	public void addColumnWidth(final int index,final int width){
		this.widths.put(index, width);
	}
	
	public void addRowHeight(final int index, final float height){
		this.heigths.put(index, height);
	}
	
	public Collection<Integer> getKeysOfWidth(){
		return this.widths.keySet();
	}
	
	public Collection<Integer> getKeysOfHeight(){
		return this.heigths.keySet();
	}
	
	/**
	 * @return the defaultWidth
	 */
	public int getDefaultWidth() {
		return defaultWidth;
	}

	/**
	 * @return the defaultHeight
	 */
	public float getDefaultHeight() {
		return defaultHeight;
	}

	/**
	 * @param defaultWidth the defaultWidth to set
	 */
	public void setDefaultWidth(final int defaultWidth) {
		this.defaultWidth = defaultWidth;
	}

	/**
	 * @param defaultHeight the defaultHeight to set
	 */
	public void setDefaultHeight(final float defaultHeight) {
		this.defaultHeight = defaultHeight;
	}

	public void setStyle(final Region region, final EStyle style){
		this.regionStyles.put(region,style);
	}

	public void setStyle(final int row, final int column, final EStyle style){
		this.setStyle(new Region(row, row, column,column), style);
	}

	public void setStyle(final int row, final EStyle style){
		this.setStyle(new Region(row, row, -1, -1), style);
	}

	public void setDefaultStyle(final EStyle style){
		this.setStyle(new Region(-1,-1,-1,-1), style);
	}

	public EStyle getRegionStyle(final Region key){
		return this.regionStyles.get(key);
	}

	public EStyle getRowTypeStyle(final RowType key){
		return this.rowTypeStyles.get(key);
	}

	public Collection<Region> regionStyleKeys(){
		return this.regionStyles.keySet();
	}

	public void addRowStyle(final RowType type, final EStyle style){
		this.rowTypeStyles.put(type, style);
	}

	public Collection<RowType> rowTypeStyleKeys(){
		return this.rowTypeStyles.keySet();
	}

	public void addRowHeight(final float height, final RowType type){
		this.heigths.put(type.getNum(), height);
	}
}
