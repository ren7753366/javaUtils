package com.changyou.common.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelUtil<T> {


	/**
	 * 功能: 导出为Excel工作簿 参数: sheetName[工作簿中的一张工作表的名称] 参数: titleName[表格的标题名称] 参数:
	 * headers[表格每一列的列名] 参数: dataSet[要导出的数据源] 参数: resultUrl[导出的excel文件地址] 参数:
	 * pattern[时间类型数据的格式]
	 */
	public static void exportExcel(String sheetName, String titleName, String[] headers, Collection<?> dataSet,
			String resultUrl, String pattern) {

		doExportExcel(sheetName, titleName, headers, dataSet, resultUrl, pattern);

	}

	/**
	 * 功能:真正实现导出
	 */
	private static void doExportExcel(String sheetName, String titleName, String[] headers, Collection<?> dataSet,
			String resultUrl, String pattern) {

		// 声明一个工作薄
		HSSFWorkbook workbook = new HSSFWorkbook();

		// 生成一个工作表
		HSSFSheet sheet = workbook.createSheet(sheetName);
		// 设置工作表默认列宽度为20个字节
		sheet.setDefaultColumnWidth((short) 20);
		// 在工作表中合并首行并居中
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, headers.length - 1));

		// 创建[标题]样式
		HSSFCellStyle titleStyle = workbook.createCellStyle();
		// 设置[标题]样式
		titleStyle.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
		titleStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		titleStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		titleStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		titleStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		titleStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		titleStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 创建[标题]字体
		HSSFFont titleFont = workbook.createFont();
		// 设置[标题]字体
		titleFont.setColor(HSSFColor.WHITE.index);
		titleFont.setFontHeightInPoints((short) 24);
		titleFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把[标题字体]应用到[标题样式]
		titleStyle.setFont(titleFont);

		// 创建[列首]样式
		HSSFCellStyle headersStyle = workbook.createCellStyle();
		// 设置[列首]样式
		headersStyle.setFillForegroundColor(HSSFColor.ORANGE.index);
		headersStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headersStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		headersStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		headersStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		headersStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		headersStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		// 创建[列首]字体
		HSSFFont headersFont = workbook.createFont();
		// 设置[列首]字体
		headersFont.setColor(HSSFColor.WHITE.index);
		headersFont.setFontHeightInPoints((short) 12);
		headersFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		// 把[列首字体]应用到[列首样式]
		headersStyle.setFont(headersFont);

		// 创建[表中数据]样式
		HSSFCellStyle dataSetStyle = workbook.createCellStyle();
		// 设置[表中数据]样式
		dataSetStyle.setFillForegroundColor(HSSFColor.WHITE.index);
		dataSetStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		dataSetStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
		dataSetStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
		dataSetStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
		dataSetStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
		dataSetStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
		dataSetStyle.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);
		// 创建[表中数据]字体
		HSSFFont dataSetFont = workbook.createFont();
		// 设置[表中数据]字体
		dataSetFont.setBoldweight(HSSFFont.BOLDWEIGHT_NORMAL);
		dataSetFont.setColor(HSSFColor.BLACK.index);
		// 把[表中数据字体]应用到[表中数据样式]
		dataSetStyle.setFont(dataSetFont);

		// 创建标题行-增加样式-赋值
		HSSFRow titleRow = sheet.createRow(0);
		if(titleName!=null){
			HSSFCell titleCell = titleRow.createCell(0);
			titleCell.setCellStyle(titleStyle);
			titleCell.setCellValue(titleName);
		}

		// 创建列首-增加样式-赋值
		HSSFRow row = sheet.createRow(1);
		for (short i = 0; i < headers.length; i++) {

			@SuppressWarnings("deprecation")
			HSSFCell cell = row.createCell(i);
			cell.setCellStyle(headersStyle);
			HSSFRichTextString text = new HSSFRichTextString(headers[i]);
			cell.setCellValue(text);

		}

		// 创建表中数据行-增加样式-赋值
		Iterator<?> it = dataSet.iterator();
		int index = 1;
		while (it.hasNext()) {
			index++;
			row = sheet.createRow(index);
			Object t = it.next();
			// 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
			Field[] fields = t.getClass().getDeclaredFields();
			for (short i = 0; i < fields.length; i++) {
				@SuppressWarnings("deprecation")
				HSSFCell cell = row.createCell(i);
				cell.setCellStyle(dataSetStyle);
				Field field = fields[i];
				String fieldName = field.getName();
				String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
				try {
					@SuppressWarnings("rawtypes")
					Class tCls = t.getClass();
					@SuppressWarnings("unchecked")
					Method getMethod = tCls.getMethod(getMethodName, new Class[] {});
					Object value = getMethod.invoke(t, new Object[] {});
					
					// 如果是时间类型,按照格式转换
					String textValue = null;
					if (value instanceof Date) {
						Date date = (Date) value;
						SimpleDateFormat sdf = new SimpleDateFormat(pattern);
						textValue = sdf.format(date);
					} else {
						// 其它数据类型都当作字符串简单处理
						textValue = value==null?"":value.toString();
					}

					// 利用正则表达式判断textValue是否全部由数字组成
					if (textValue != null) {
						Pattern p = Pattern.compile("^\\d+(\\.\\d+)?$");
						Matcher matcher = p.matcher(textValue);
						if (matcher.matches()) {
							// 是数字当作double处理
							cell.setCellValue(Double.parseDouble(textValue));
						} else {
							// 不是数字做普通处理
							cell.setCellValue(textValue);
						}
					}

					OutputStream out = null;
					try {
						out = new FileOutputStream(resultUrl);
						workbook.write(out);
					} catch (IOException e) {
						e.printStackTrace();
					} finally {
						try {
							out.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}

				} catch (SecurityException e) {
					e.printStackTrace();
				} catch (NoSuchMethodException e) {
					e.printStackTrace();
				} catch (IllegalArgumentException e) {
					e.printStackTrace();
				} catch (IllegalAccessException e) {
					e.printStackTrace();
				} catch (InvocationTargetException e) {
					e.printStackTrace();
				} finally {
					// 清理资源
					try {
						workbook.close();
					} catch (IOException e) {
						e.printStackTrace();
					}
				}
			}
		}
	}
	
	//正则表达式 用于匹配属性的第一个字母
    private static final String REGEX = "[a-zA-Z]";

    /**
     * 功能: Excel数据导入到数据库
     * 参数: originUrl[Excel表的所在路径]
     * 参数: startRow[从第几行开始]
     * 参数: endRow[到第几行结束
     *                  (0表示所有行;
     *                  正数表示到第几行结束;
     *                  负数表示到倒数第几行结束)]
     * 参数: clazz[要返回的对象集合的类型]
     */
    public static List<?> importExcel(String originUrl,int startRow,int endRow,Class<?> clazz) throws IOException {
        //是否打印提示信息
        boolean showInfo=true;
        return doImportExcel(originUrl,startRow,endRow,showInfo,clazz);
    }

    /**
     * 功能:真正实现导入
     */
    private static List<Object> doImportExcel(String originUrl,int startRow,int endRow,boolean showInfo,Class<?> clazz) throws IOException {
        // 判断文件是否存在
        File file = new File(originUrl);
        if (!file.exists()) {
            throw new IOException("文件名为" + file.getName() + "Excel文件不存在！");
        }
        HSSFWorkbook wb = null;
        FileInputStream fis=null;
        List<Row> rowList = new ArrayList<Row>();
        try {
            fis = new FileInputStream(file);
            // 去读Excel
            wb = new HSSFWorkbook(fis);
            Sheet sheet = wb.getSheetAt(0);
            // 获取最后行号
            int lastRowNum = sheet.getLastRowNum();
            if (lastRowNum > 0) { // 如果>0，表示有数据
                out("\n开始读取名为【" + sheet.getSheetName() + "】的内容：",showInfo);
            }
            Row row = null;
            // 循环读取
            for (int i = startRow; i <= lastRowNum + endRow; i++) {
                row = sheet.getRow(i);
                if (row != null) {
                    rowList.add(row);
                    out("第" + (i + 1) + "行：",showInfo,false);
                    // 获取每一单元格的值
                    for (int j = 0; j < row.getLastCellNum(); j++) {
                        String value = getCellValue(row.getCell(j));
                        if (!value.equals("")) {
                            out(value + " | ",showInfo,false);
                        }
                    }
                    out("",showInfo);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally{
            wb.close();
        }
        return returnObjectList(rowList,clazz);
    }

    /**
     * 功能:获取单元格的值
     */
    private static String getCellValue(Cell cell) {
        Object result = "";
        if (cell != null) {
            switch (cell.getCellType()) {
            case Cell.CELL_TYPE_STRING:
                result = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_NUMERIC:
                result = cell.getNumericCellValue();
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                result = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_FORMULA:
                result = cell.getCellFormula();
                break;
            case Cell.CELL_TYPE_ERROR:
                result = cell.getErrorCellValue();
                break;
            case Cell.CELL_TYPE_BLANK:
                break;
            default:
                break;
            }
        }
        return result.toString();
    }

    /**
     * 功能:返回指定的对象集合
     */
    private static List<Object> returnObjectList(List<Row> rowList,Class<?> clazz) {
        List<Object> objectList=null;
        Object obj=null;
        String attribute=null;
        String value=null;
        int j=0;
        try {   
            objectList=new ArrayList<Object>();
            Field[] declaredFields = clazz.getDeclaredFields();
            for (Row row : rowList) {
                j=0;
                obj = clazz.newInstance();
                for (Field field : declaredFields) {    
                    attribute=field.getName().toString();
                    value = getCellValue(row.getCell(j));
                    setAttrributeValue(obj,attribute,value);    
                    j++;
                }
                objectList.add(obj);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return objectList;
    }

    /**
     * 功能:给指定对象的指定属性赋值
     */
    private static void setAttrributeValue(Object obj,String attribute,String value) {
        //得到该属性的set方法名
        String method_name = convertToMethodName(attribute,obj.getClass(),true);
        Method[] methods = obj.getClass().getMethods();
        for (Method method : methods) {
            /**
             * 因为这里只是调用bean中属性的set方法，属性名称不能重复
             * 所以set方法也不会重复，所以就直接用方法名称去锁定一个方法
             * （注：在java中，锁定一个方法的条件是方法名及参数）
            */
            if(method.getName().equals(method_name))
            {
                Class<?>[] parameterC = method.getParameterTypes();
                try {
                    /**如果是(整型,浮点型,布尔型,字节型,时间类型),
                     * 按照各自的规则把value值转换成各自的类型
                     * 否则一律按类型强制转换(比如:String类型)
                    */
                    if(parameterC[0] == int.class || parameterC[0]==java.lang.Integer.class)
                    {
                        value = value.substring(0, value.lastIndexOf("."));
                        method.invoke(obj,Integer.valueOf(value));
                        break;
                    }else if(parameterC[0] == float.class || parameterC[0]==java.lang.Float.class){
                        method.invoke(obj, Float.valueOf(value));
                        break;
                    }else if(parameterC[0] == double.class || parameterC[0]==java.lang.Double.class)
                    {
                        method.invoke(obj, Double.valueOf(value));
                        break;
                    }else if(parameterC[0] == byte.class || parameterC[0]==java.lang.Byte.class)
                    {
                        method.invoke(obj, Byte.valueOf(value));
                        break;
                    }else if(parameterC[0] == boolean.class|| parameterC[0]==java.lang.Boolean.class)
                    {
                        method.invoke(obj, Boolean.valueOf(value));
                        break;
                    }else if(parameterC[0] == java.util.Date.class)
                    {
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        Date date=null;
                        try {
                            date=sdf.parse(value);
                        } catch (Exception e) {
                            e.printStackTrace();
                        }
                        method.invoke(obj,date);
                        break;
                    }else
                    {
                        method.invoke(obj,parameterC[0].cast(value));
                        break;
                    }
                } catch (IllegalArgumentException e) {
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    e.printStackTrace();
                } catch (SecurityException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    /**
     * 功能:根据属性生成对应的set/get方法
     */
    private static String convertToMethodName(String attribute,Class<?> objClass,boolean isSet) {
        /** 通过正则表达式来匹配第一个字符 **/
        Pattern p = Pattern.compile(REGEX);
        Matcher m = p.matcher(attribute);
        StringBuilder sb = new StringBuilder();
        /** 如果是set方法名称 **/
        if(isSet)
        {
            sb.append("set");
        }else{
        /** get方法名称 **/
            try {
                Field attributeField = objClass.getDeclaredField(attribute);
                /** 如果类型为boolean **/
                if(attributeField.getType() == boolean.class||attributeField.getType() == Boolean.class)
                {
                    sb.append("is");
                }else
                {
                    sb.append("get");
                }
            } catch (SecurityException e) {
                e.printStackTrace();
            } catch (NoSuchFieldException e) {
                e.printStackTrace();
            }
        }
        /** 针对以下划线开头的属性 **/
        if(attribute.charAt(0)!='_' && m.find())
        {
            sb.append(m.replaceFirst(m.group().toUpperCase()));
        }else{
            sb.append(attribute);
        }
        return sb.toString();
    }

    /**
     * 功能:输出提示信息(普通信息打印)
     */
    private static void out(String info, boolean showInfo) {
        if (showInfo) {
            System.out.print(info + (showInfo ? "\n" : ""));
        }
    }

    /**
     * 功能:输出提示信息(同一行的不同单元格信息打印)
     */
    private static void out(String info, boolean showInfo, boolean nextLine) {
        if (showInfo) {
            if(nextLine)
            {
                System.out.print(info + (showInfo ? "\n" : ""));
            }else
            {
                System.out.print( info );
            }
        }
    }
    
    static class Book{
    	private int id;
    	private String name;
    	private double price;
    	private Date time;
		public int getId() {
			return id;
		}
		public void setId(int id) {
			this.id = id;
		}
		public String getName() {
			return name;
		}
		public void setName(String name) {
			this.name = name;
		}
		public double getPrice() {
			return price;
		}
		public void setPrice(double price) {
			this.price = price;
		}
		public Date getTime() {
			return time;
		}
		public void setTime(Date time) {
			this.time = time;
		}
		public Book(int id, String name, double price, Date time) {
			super();
			this.id = id;
			this.name = name;
			this.price = price;
			this.time = time;
		}
    }

    public static void main(String[] args) {
    	String sheetName="图书列表";
//        String titleName="我的图书";
    	String titleName=null;
        String[] headers = { "图书id", "图书名称", "图书价格", "入库时间" };
        List<Book> dataSet = new ArrayList<Book>();
        dataSet.add(new Book(1,"十万个为什么",7.36,new Date()));
        dataSet.add(new Book(2,"十万个为什么1",7.36,new Date()));
        dataSet.add(new Book(3,"十万个为什么2",7.36,new Date()));
        String resultUrl="E:\\book.xls";
        String pattern="yyyy-MM-dd HH:mm";
        exportExcel(sheetName, titleName, headers, dataSet, resultUrl, pattern);
	}
}
