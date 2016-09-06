/**
 * 
 */
package template;

import xgt.util.excel.Pager;
import xgt.util.excel.model.Record;

import java.util.Collection;

class MyPager implements Pager{

	/**
	 * @param title
	 * @param headers
	 * @param body
	 */
	public MyPager(String title, Record header, Collection<? extends Record> body) {
		super();
		this.title = title;
		this.headers = header;
		this.body = body;
	}

	private String title;
	
	private Record headers;
	
	private Collection<? extends Record> body;
	
	public String getTitle() {
		return title;
	}

	public int getPageNum() {
		return 0;
	}

	public int getPageColumn(){
		try {
			return headers.getCells().length;
		} catch (IllegalAccessException e) {
			return 0;
		}
	}

	public Record getHeader() {
		return headers;
	}

	public Collection<? extends Record> getBody() {
		return body;
	}

}