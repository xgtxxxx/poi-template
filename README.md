Base on Poi

1. Read(){
  final Reader reader = SimpleReaderTemplate.newInstance();
  final InputStream is = new FileInputStream(new File("D:/text.xls"));
  final List<ESheet> sheets = reader.read(is);
  for(final ESheet sheet: sheets){
      for(final ERow row: sheet.getRows()){
          for(final ECell cell: row.getCells()){
              System.out.print(cell);
              System.out.println("----"+cell.getValue().getClass());
          }
      }
  }
}

2. write example1: 导出简单的excel，并且自己设置样式
public static void main(String[] args) {
		final Test1Model header = new Test1Model(new Object[]{"周一","周二","周三","周四","周五","周六","周日"});
		final List<Test1Model> body = new ArrayList<Test1Model>();
		for(int i=0; i<6; i++){
			final Object[] data = new Object[7];
			data[1] = "String"+i;
			data[2] = new Object();
			data[3] = 10+i;
			data[4] = 20.5+i;
			data[5] = true;
			data[6] = new Date();
			data[0] = null;
			body.add(new Test1Model(data));
		}
		final EasyTemplate t = TemplateFactory.createTemplate(SimpleTemplate.class, header, body);
		final Config config = t.getSheets().get(0).getSheetStyle();
		EStyle style = EStyle.newBorderInstance();
		StyleDecorate.decorateBgYellow(style);
		style.setFontSize(EStyle.FONT_SIZE_BIG);
		style.setFontBold(true);
		style.setAlign(CellStyle.ALIGN_CENTER);
		config.setStyle(new Region(0,0,0,6),style);

		style = EStyle.newBorderInstance();
		style.setFormat(EStyle.FORMAT_MONEY_RMB);
		config.setStyle(new Region(1,6,4,4), style);

		style = EStyle.newBorderInstance();
		style.setFormat(EStyle.FORMAT_PERCENTAGE);
		config.setStyle(new Region(1,6,3,3),style);

		style = EStyle.newBorderInstance();
		style.setFormat(EStyle.FORMAT_DATE);
		config.setStyle(new Region(1,6,6,6), style);

		config.addRowHeight(0, 30f);
		config.addColumnWidth(2, Config.DEFAULT_WIDTH*3);
		config.addColumnWidth(6, Config.DEFAULT_WIDTH+3);
		
		FileOutputStream fos = null;
		try {
			fos = new FileOutputStream("D:/text.xlsx");
			t.build(fos);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
3. example: 导出一页Page
public static void main(String[] args) {
	String[] header = { "周一", "周二", "周三", "周四", "周五", "周六", "周日" };
      Test1Model h = new Test1Model(header);
	List<Test1Model> list = new ArrayList<Test1Model>();
	for (int i = 0; i < 6; i++) {
		Object[] data = new Object[7];
		data[1] = "String" + i;
		data[2] = new Object();
		data[3] = 10 + i;
		data[4] = 20.5 + i;
		data[5] = true;
		data[6] = new Date();
		data[0] = null;
		list.add(new Test1Model(data));
	}

	Pager pager = new MyPager("测试", h, list);

	EasyTemplate t = TemplateFactory.createTemplate(SimpleTemplate.class,pager);

	Config config = t.getConfig(0);
	config.setDefaultWidth(Config.DEFAULT_WIDTH+5);

	FileOutputStream fos = null;
	try {
		fos = new FileOutputStream("D:/text.xlsx");
		t.build(fos);
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	} finally {
		try {
			if (fos != null) {
				fos.close();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

4.导出多页
public static void main(String[] args) {
	Pagers pagers = new Pagers();
	for (int j = 0; j < 5; j++) {
		Record header = new Test1Model(new Object[]{ "周一", "周二", "周三", "周四", "周五", "周六", "周日" });

		Collection<Record> list = new ArrayList<Record>();
		for (int i = 0; i < 6; i++) {
			Object[] data = new Object[7];
			data[1] = "String" + i;
			data[2] = new Object();
			data[3] = 10 + i;
			data[4] = 20.5 + i;
			data[5] = true;
			data[6] = new Date();
			data[0] = null;
			list.add(new Test1Model(data));
		}
		MyPager pager = new MyPager("测试" + j, header, list);
		pagers.addPager(pager);
	}

	pagers.setTitleOnlyFirstPage(false);
	pagers.setLineSpacing(1);

	EasyTemplate t = TemplateFactory.createTemplate(SimpleTemplate.class,pagers);

	Config config = t.getConfig(0);

	config.setDefaultWidth(config.getDefaultWidth()*2);
      config.addRowHeight(config.getDefaultHeight()+10, RowType.TITLE);

	FileOutputStream fos = null;
	try {
		fos = new FileOutputStream("D:/text.xls");
		t.build(fos);
	} catch (FileNotFoundException e) {
		e.printStackTrace();
	} catch (IOException e) {
		e.printStackTrace();
	} finally {
		try {
			if (fos != null) {
				fos.close();
			}
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

5. 导出多个sheet
public static void main(final String[] args) throws IllegalAccessException, IOException {
    final EasyTemplate template = new SimpleTemplate();
    for(int j=0; j<3; j++){
        final SheetTemplate sheet = new SheetTemplate();
        sheet.setName("name"+j);
        final Record header = new Test1Model(new Object[]{ "周一", "周二", "周三", "周四", "周五", "周六", "周日" });
        final Collection<Record> list = new ArrayList<Record>();
        for (int i = 0; i < 6; i++) {
            final Object[] data = new Object[7];
            data[1] = "String" + i;
            data[2] = new Object();
            data[3] = 10 + i;
            data[4] = 20.5 + i;
            data[5] = true;
            data[6] = new Date();
            data[0] = null;
            list.add(new Test1Model(data));
        }
        sheet.addData(header, list);
        template.addSheet(sheet);
    }

    final FileOutputStream fos = new FileOutputStream("D:/text.xls");
    template.build(fos);
}
	
