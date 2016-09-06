package xgt.util.excel;

import xgt.util.excel.model.ESheet;

import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface Reader {
    public List<ESheet> read(final InputStream is) throws IOException;
}
