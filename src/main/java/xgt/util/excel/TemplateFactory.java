/**
 * 
 */
package xgt.util.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.log4j.Logger;
import xgt.util.excel.model.Pagers;
import xgt.util.excel.model.Record;

import java.util.Collection;
import java.util.List;

/**
 * @author Gavin
 *
 */
public final class TemplateFactory {
	
	private static final Logger LOG = Logger.getLogger(TemplateFactory.class);

	public static EasyTemplate createTemplate(final Class<? extends EasyTemplate> clazz, final Record header, final Collection<? extends Record> datas){
		EasyTemplate t = null;
		try {
			t = clazz.newInstance();
			t.addSheet(createSheetTemplate(header, datas));
		} catch (final IllegalAccessException e) {
			LOG.error(e);
		} catch (final InstantiationException e) {
			LOG.error(e);
		}
		return t;
	}

	private static SheetTemplate createSheetTemplate(final Record header, final Collection<? extends Record> datas) throws IllegalAccessException {
		final SheetTemplate st = new SheetTemplate();
		st.addData(header, datas);
		return st;
	}

	public static EasyTemplate createTemplate(final Class<? extends EasyTemplate> clazz, final Pagers pagers){
		EasyTemplate t = null;
		try {
			t = clazz.newInstance();
			final SheetTemplate sheetTemplate = new SheetTemplate();
			final List<Pager> ps = pagers.getPagers();
			for (int i = 0; i<ps.size(); i++) {
				final Pager pager = ps.get(i);
				if(pagers.isTitleOnlyFirstPage()){
					if(i==0){
						sheetTemplate.addTitle(pager.getTitle(), pager.getPageColumn());
					}
				}else{
					sheetTemplate.addTitle(pager.getTitle(), pager.getPageColumn());
				}
				sheetTemplate.addData(pager.getHeader(), pager.getBody());
				sheetTemplate.addBlank(pagers.getLineSpacing());
			}
			t.addSheet(sheetTemplate);
		} catch (final IllegalAccessException e) {
			LOG.error(e);
		} catch (final InstantiationException e) {
			LOG.error(e);
		}
		return t;
	}
	
	public static EasyTemplate createTemplate(final Class<? extends EasyTemplate> clazz, final Pager pager){
		EasyTemplate t = null;
		try {
			t = clazz.newInstance();
			final SheetTemplate st = new SheetTemplate();
			if(StringUtils.isNotEmpty(pager.getTitle())){
				st.addTitle(pager.getTitle(), pager.getPageColumn(), 0);
			}
			st.addData(pager.getHeader(), pager.getBody());
			t.addSheet(st);
		} catch (final IllegalAccessException e) {
			LOG.error(e);
		} catch (final InstantiationException e) {
			LOG.error(e);
		}
		return t;
	}
}
