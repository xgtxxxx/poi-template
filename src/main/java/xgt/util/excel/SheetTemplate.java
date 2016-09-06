package xgt.util.excel;

import xgt.util.excel.model.ERow;
import xgt.util.excel.model.ESheet;
import xgt.util.excel.model.Record;
import xgt.util.excel.model.RowType;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

public class SheetTemplate {
    private ESheet sheetData = new ESheet();
    private List<Region> regions = new ArrayList<Region>();
    private Config sheetStyle = new Config();

    public ESheet getSheetData() {
        return sheetData;
    }

    public void setSheetData(final ESheet sheetData) {
        this.sheetData = sheetData;
    }

    public List<Region> getRegions() {
        return regions;
    }

    public void setRegions(final List<Region> regions) {
        this.regions = regions;
    }

    public Config getSheetStyle() {
        return sheetStyle;
    }

    public void setSheetStyle(final Config sheetStyle) {
        this.sheetStyle = sheetStyle;
    }

    public SheetTemplate newSimpleInstance(final Record header, final Collection<Record> body) throws IllegalAccessException {
        final SheetTemplate st = new SheetTemplate();
        st.addData(header, body);
        return st;
    }

    public int addTitle(final String title,final int columnspan){
        final int index = this.sheetData.getRows().size();
        return this.addTitle(title, columnspan, index);
    }

    public int addTitle(final String title,final int columnspan, final int index){
        this.sheetData.addRow(ERow.newTitleInstance(title, index));
        this.regions.add(new Region(index, index, 0, columnspan-1));
        return index+1;
    }

    public int addData(final Record record, final RowType rowType) throws IllegalAccessException {
        this.sheetData.addRow(ERow.newInstance(record.getCells(), this.sheetData.getRows().size(), rowType));
        return this.sheetData.getRows().size();
    }

    public int addData(final Record header, final Collection<? extends Record> body) throws IllegalAccessException {
        this.addHeader(header);
        this.addBody(body);
        return this.sheetData.getRows().size();
    }

    public int addHeader(final Record header, final int index) throws IllegalAccessException {
        this.sheetData.addRow(ERow.newHeaderInstance(header.getCells(), index));
        return this.sheetData.getRows().size();
    }

    public int addHeader(final Record header) throws IllegalAccessException {
        this.sheetData.addRow(ERow.newHeaderInstance(header.getCells(), this.sheetData.getRows().size()));
        return this.sheetData.getRows().size();
    }

    public int addBody(final Collection<? extends Record> body) throws IllegalAccessException {
        int startRow = this.sheetData.getRows().size();
        for(final Record record : body){
            this.sheetData.addRow(ERow.newBodyInstance(record.getCells(), startRow));
            startRow = startRow+1;
        }
        return this.sheetData.getRows().size();
    }

    public int addBlank(final int rows){
        int index = this.sheetData.getRows().size();
        for(int i=0; i<rows; i++){
            this.sheetData.addRow(ERow.newBlankInstance(index));
            index = index+1;
        }
        return index;
    }

    public void setName(final String name){
        this.sheetData.setName(name);
    }
}
