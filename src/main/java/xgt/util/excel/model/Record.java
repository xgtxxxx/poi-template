package xgt.util.excel.model;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public abstract class Record {

    @Ignore
    private Object defaultValueIfNull;

    public Object[] getCells() throws IllegalAccessException {
        final List<Object> objs = new ArrayList<Object>();
        final Field[] fields = this.getClass().getDeclaredFields();
        for(final Field field: fields){
            field.setAccessible(true);
            final Ignore ignore = field.getAnnotation(Ignore.class);
            if(ignore==null){
                Object value = field.get(this);
                if(value==null){
                    value = defaultValueIfNull;
                }
                objs.add(value);
            }
        }
        return objs.toArray();
    }

    public void setDefaultValueIfNull(final Object defaultValueIfNull) {
        this.defaultValueIfNull = defaultValueIfNull;
    }
}
