/**
 * 
 */
package xgt.util.excel.model;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import xgt.util.excel.Pager;

/**
 * @author Gavin
 *
 */
public class Pagers implements Iterable<Pager>{
	
	private boolean titleOnlyFirstPage;
	
	private int lineSpacing;
	
	private List<Pager> pagers;
	
	public int addPager(final Pager pager){
		if(pagers == null){
			pagers = new ArrayList<Pager>();
		}
		pagers.add(pager);
		return pagers.size();
	}
	
	public List<Pager> getPagers(){
		return pagers==null?new ArrayList<Pager>():pagers;
	}
	
	public int getPagerSize(){
		return pagers==null?0:pagers.size();
	}

	public Iterator<Pager> iterator() {
		if(pagers == null){
			pagers = new ArrayList<Pager>();
		}
		return pagers.iterator();
	}

	/**
	 * @return the titleOnlyFirstPage
	 */
	public boolean isTitleOnlyFirstPage() {
		return titleOnlyFirstPage;
	}

	/**
	 * @param titleOnlyFirstPage the titleOnlyFirstPage to set
	 */
	public void setTitleOnlyFirstPage(boolean titleOnlyFirstPage) {
		this.titleOnlyFirstPage = titleOnlyFirstPage;
	}

	/**
	 * @return the lineSpacing
	 */
	public int getLineSpacing() {
		return lineSpacing;
	}

	/**
	 * @param lineSpacing the lineSpacing to set
	 */
	public void setLineSpacing(int lineSpacing) {
		this.lineSpacing = lineSpacing;
	}
	
}
