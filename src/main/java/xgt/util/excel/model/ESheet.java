package xgt.util.excel.model;

import java.util.ArrayList;
import java.util.List;

public class ESheet {
    private String name;
    private List<ERow> rows = new ArrayList<ERow>();

    public ESheet(){}

    public ESheet(final String name){
        this.name = name;
    }

    public void addRow(final ERow row){
        rows.add(row);
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public List<ERow> getRows() {
        return rows;
    }

    public void setRows(List<ERow> rows) {
        this.rows = rows;
    }
}
