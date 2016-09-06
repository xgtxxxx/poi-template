package xgt.util.excel.model;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public abstract class Record {
    public Object[] getCells() throws IllegalAccessException {
        final List<Object> objs = new ArrayList<Object>();
        final Field[] fields = this.getClass().getDeclaredFields();
        for(final Field field: fields){
            field.setAccessible(true);
            final Ignore ignore = field.getAnnotation(Ignore.class);
            if(ignore==null){
                objs.add(field.get(this));
            }
        }
        return objs.toArray();
    }
}
