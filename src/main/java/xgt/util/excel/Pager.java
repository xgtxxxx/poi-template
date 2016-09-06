/**
 * 
 */
package xgt.util.excel;

import xgt.util.excel.model.Record;

import java.util.Collection;

/**
 * @author Gavin
 *
 */
public interface Pager {
	public String getTitle();
	
	public int getPageNum();

	public int getPageColumn();
	
	public Record getHeader();
	
	public Collection<? extends Record> getBody();
	
}
