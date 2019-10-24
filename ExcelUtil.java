package com.bgyfw.comprehensive.util;

import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.context.request.RequestAttributes;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Comparator;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static java.util.stream.Collectors.toList;

@Slf4j
public class ExcelUtil {

    /**
     *  导入excel文件,返回数据列表
     * @param mfile
     * @param clazz
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importData(MultipartFile mfile, Class<T> clazz) throws Exception {

        String originalFilename = mfile.getOriginalFilename();
        String fileSuffix = originalFilename.substring(originalFilename.lastIndexOf(".")+1, originalFilename.length());
        boolean contains = Arrays.asList(new String[]{"xls","xlsx","xlsb","xlsm"}).contains(fileSuffix);
        if (contains==false) {
        }

        InputStream stream = mfile.getInputStream();
        // 判断文件格式
        boolean isExcel03 = mfile.getOriginalFilename().matches(".+\\.(?i)(xls)");//这里的(?i)代表忽略大小写
        Workbook wb = isExcel03 ? new HSSFWorkbook(stream) : new XSSFWorkbook(stream);
        // 读取工作表
        Sheet sheet = wb.getSheetAt(0);    // sheet1
        List<T> resultData = new ArrayList<>();    //存放返回数据
        // 开始每行的读取
        //System.out.println(sheet.getPhysicalNumberOfRows());
        Row firstRow = sheet.getRow(0);

        HashMap<Integer, String> fieldMap = new HashMap<>();
        short titlerow = firstRow.getLastCellNum();
        for (int i = 0; i < titlerow; i++) {
            Cell cell = firstRow.getCell(i);
            Field[] declaredFields = clazz.getDeclaredFields();
            //判断标题列是否正确
            boolean flag = false;
            for (Field field : declaredFields) {

                String cellValue = "";
                if (null != cell) {
                    cellValue = cell.getStringCellValue();
                    cellValue = cellValue.trim();
                    ExcelField excelField = field.getAnnotation(ExcelField.class);
                    if (excelField != null && cellValue.equals(excelField.name().trim())) {
                        fieldMap.put(i, field.getName());
                        flag = true;
                    }
                }
            }
            if (!flag) {
            }
        }


        int maxRowLength = 0;
        if (0 != sheet.getLastRowNum()) {
            if (0 != sheet.getRow(0).getLastCellNum()) {
                maxRowLength = sheet.getRow(0).getLastCellNum();
            }
        }
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            boolean flag = true;
            Row row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            T t = clazz.newInstance();
            for (int j = 0; j < maxRowLength; j++) {    // for循环可读取 null 值
                Cell cell = row.getCell(j);
                Object cellValue = "";
                if (null != cell) {
                    switch (cell.getCellType()) {
                        case BLANK:    // 空白  //Cell.CELL_TYPE_STRING:		// 文本
                            cellValue = "";
                            break;
                        case STRING:
                            cellValue = cell.getStringCellValue();
                            break;
                        case NUMERIC: //Cell.CELL_TYPE_NUMERIC:	// 数字或者日期
                            if (DateUtil.isCellDateFormatted(cell)) {        // 是否是日期
                                cellValue = cell.getDateCellValue();
                            } else {
                                cell.setCellType(CellType.STRING);
                                cellValue = cell.getStringCellValue();
                            }
                            break;
                        case BOOLEAN: //Cell.CELL_TYPE_BOOLEAN:	// 布尔型
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        default:
                            cellValue = null;
                            break;
                    }
                }

                String fieldName = fieldMap.get(j);
                if (StringUtils.isNotBlank(fieldName)) {
                    Field field = clazz.getDeclaredField(fieldName);
                    field.setAccessible(true);

                    try {
                        log.debug("field value {}", cellValue);
                        log.debug("fieldName value {}", fieldName);
                        if (!"".equals(cellValue) && null != cellValue) {
                            flag = false;
                        }
                        field.set(t, cellValue);
                    } catch (IllegalArgumentException e) {
                        Class<?> type = field.getType();
                        if (type == BigDecimal.class) {
                            if (cellValue!=null && StringUtils.isNotBlank(cellValue.toString())) {
                                BigDecimal b = new BigDecimal(cellValue.toString());
                                field.set(t, b);
                            }
                        } else if (type == Integer.class ||type == int.class) {
                            if (cellValue!=null && StringUtils.isNotBlank(cellValue.toString())) {
                                field.set(t, Integer.parseInt(cellValue.toString()));
                            }
                        } else if (type == Double.class ||type == double.class) {
                            if (cellValue!=null && StringUtils.isNotBlank(cellValue.toString())) {
                                field.set(t, Double.parseDouble(cellValue.toString()));
                            }
                        }

                    } catch (IllegalAccessException e) {
                    }
                }


            }

            if (!flag) {
                resultData.add(t);    // 每行数据存入返回数据集合
            }
        }

        // 关流
        wb.close();
        stream.close();

        return resultData;
    }


    /**
     * 校验数据
     * @param <T>
     * @param list
     * @return
     * @throws NoSuchMethodException
     * @throws InvocationTargetException
     * @throws IllegalAccessException
     * @throws InstantiationException
     */
    public static <T> Map<String, List<T>> validate(List<T> list, IValidator validator) throws Exception {

        validator.beforeValidate();
        Class<? extends IValidator> validateClass = validator.getClass();

        HashMap<String, List<T>> map = new HashMap<>();
        ArrayList<T> success = new ArrayList<>();
        ArrayList<T> error = new ArrayList<>();

        //遍历数据列表
        for (T t : list) {
            Class<?> aClass = t.getClass();
            Field[] declaredFields = aClass.getDeclaredFields();
            boolean flag =true;

            //查出有ExcelField注解的字段,进行校验
            for (Field field : declaredFields) {

                ExcelField annotation = field.getAnnotation(ExcelField.class);
                if (annotation != null) {

                    //注解上标注的验证方法名
                    String[] methods = annotation.validateMethod();
                    for (String method : methods) {
                        if (StringUtils.isNotBlank(method)) {
                            Method method1 = null;
                            Method[] methods1 = validateClass.getMethods();
                            for (Method method2 : methods1) {
                                if (method2.getName().equals(method)) {
                                    method1 = method2;
                                }
                            }

                            method1.setAccessible(true);
                            int parameterCount = method1.getParameterCount();
                            boolean invoke = false;
                            if (parameterCount == 1) {
                                invoke = (boolean) method1.invoke(validator, t);
                            } else if (parameterCount == 2) {
                                invoke = (boolean) method1.invoke(validator, t,success);
                            }else if (parameterCount == 3) {
                                invoke = (boolean) method1.invoke(validator, t,success,error);
                            }

                            if(!invoke){
                                flag = invoke;
                            }


                        }
                    }
                }

            }

            if (flag) {
                success.add(t);
            }else {
                error.add(t);
            }
            map.put("success", success);
            map.put("error", error);

        }

        validator.afterValidate(success,error);
        return map;
    }


    private static String suffix = ".xls";


    public void setSuffix(String suffix) {
        ExcelUtil.suffix = suffix;
    }








    /**
     *
     * @param dataList 数据列表
     * @param fileName 导出文件名
     * @param title excel第一行标题
     * @param clz dto类
     * @param response
     * @param required 是否导出所有字段
     * @param <T>
     */
    public static <T> void exportExcel(List<T> dataList, String fileName,
                                       String title, Class clz, HttpServletResponse response,boolean required) {
        try {
            Workbook wb = new HSSFWorkbook();
            wb = createExcel(wb, dataList, title, clz,required);
            ExcelUtil.downExcel(fileName, wb, response);

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, String fontName, short fontSize,
                                                 boolean isBlod) {
        // 通过工作簿创建样式
        HSSFCellStyle cellStyle = workbook.createCellStyle();
        // 设置水平和垂直居中
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        // 通过工作簿创建字体
        HSSFFont font = workbook.createFont();
        // 设置字体
        if (null != fontName && !"".equals(fontName)) {
            font.setFontName(fontName);
        }
        // 设置字体大小
        font.setFontHeightInPoints(fontSize);
        // 设置字体是否加粗
        font.setBold(isBlod);
        // 把字体set到样式中
        cellStyle.setFont(font);

        return cellStyle;
    }

    /**
     * @param book   工作簿
     * @param list   需写入数据
     * @return
     * @throws Exception
     */
    public static <T> HSSFWorkbook createExcel(Workbook book, List<T> list, String title, Class clz,boolean isExport) throws Exception {
        // 声明一个工作薄
        HSSFWorkbook wb = (HSSFWorkbook) book;
        //声明一个单子并命名
        wb.createSheet();

        HSSFSheet sheet = wb.getSheetAt(0);
        //给单子名称一个长度
        sheet.setDefaultColumnWidth(15);
        //给表头第一行一次创建单元格
        HSSFCell cell;
        //创建一个字体样式
        HSSFFont font = wb.createFont();
        font.setColor(IndexedColors.WHITE.getIndex());
        //向单元格里填充数据
        // 生成一个样式
        HSSFCellStyle style = wb.createCellStyle();
        HSSFCellStyle cellStyle = wb.createCellStyle();

        //默认样式字体居中
        style.setAlignment(HorizontalAlignment.CENTER);

        // 获取实体类的所有属性，返回Field数组
        Field[] fields = clz.getDeclaredFields();


        List<ExcelField> column = new ArrayList<>();
        for (int i = 0; i < fields.length; i++) {
            Field field = fields[i];
            ExcelField annotation = field.getAnnotation(ExcelField.class);
            if (annotation != null) {
                if (isExport) {
                    column.add(annotation );
                }else{
                    if (annotation.isExport()) {
                        column.add(annotation);
                    }
                }
            }
        }

        //对column排序根据order排序
        column = column.stream().sorted(Comparator.comparingInt(ExcelField::order)).collect(toList());

        //设置下拉列表
        int index =0;
        for (ExcelField value : column) {

            String[] option = value.option();
            if (option.length > 0) {
                DVConstraint constraint = DVConstraint.createExplicitListConstraint(option);
                // 设置数据有效性加载在哪个单元格上,四个参数分别是：起始行、终止行、起始列、终止列
                CellRangeAddressList regions = new CellRangeAddressList(0, 500, index, index);
                // 数据有效性对象
                HSSFDataValidation dataValidationList = new HSSFDataValidation(regions, constraint);

                sheet.addValidationData(dataValidationList);

            }
            index++;
        }

        //设置title
        HSSFRow sheetRow = null;
        if (StringUtils.isNotBlank(title)) {

            // 合并单元格
            if (column.size() == 1) {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, column.size()));
            } else {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, column.size() - 1));
            }

            /** 第一行标题处理  */
            HSSFRow row01 = sheet.createRow(0);
            HSSFCell cell01 = row01.createCell(0);
            // 设置单元格内容
            cell01.setCellValue(title);
            // 创建单元格样式
            HSSFCellStyle cellStyle01 = createCellStyle(wb, "Courier New", (short) 12, true);
            cellStyle01.setBorderBottom(BorderStyle.THIN);
            cellStyle01.setBorderLeft(BorderStyle.THIN);
            cellStyle01.setBorderRight(BorderStyle.THIN);
            cellStyle01.setBorderTop(BorderStyle.THIN);
            // 设置单元格样式
            cell01.setCellStyle(cellStyle01);
            int lastRowNum = sheet.getLastRowNum();

            sheetRow = sheet.createRow(lastRowNum + 1);
        } else {
            sheetRow = sheet.createRow(0);
        }


        //设置列名
        for (int i =0;i<column.size();i++) {

            HSSFCell cell1 = sheetRow.createCell(i);
            cell1.setCellValue(column.get(i).name());
            HSSFCellStyle style1 = wb.createCellStyle();
            style1.setBorderBottom(BorderStyle.THIN);
            style1.setBorderLeft(BorderStyle.THIN);
            style1.setBorderRight(BorderStyle.THIN);
            style1.setBorderTop(BorderStyle.THIN);
            style1.setAlignment(HorizontalAlignment.CENTER);
            style1.setVerticalAlignment(VerticalAlignment.CENTER);

            Font fontTitle = wb.createFont();
            fontTitle.setBold(true);
            style1.setFont(fontTitle);
            cell1.setCellStyle(style1);
        }




        int rowNum = sheet.getLastRowNum();
        for (int i = 0; i < list.size(); i++) {
            Object object = list.get(i);

            HSSFRow row = sheet.createRow(rowNum + i + 1);
            for (int j = 0; j < fields.length; j++) {
                Field field = fields[j];
                ExcelField annotation = field.getAnnotation(ExcelField.class);
                if (annotation == null) {
                    continue;
                }
                if (!isExport && !annotation.isExport()) {
                    continue;
                }

                int i1 = column.indexOf(annotation);

                cell = row.createCell(i1);

                cell.setCellStyle(style);
                // 如果类型是String
                if ("class java.lang.String".equals(field.getGenericType().toString())) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    String value = (String) m.invoke(object);
                    if (null == value) {
                        cell.setCellValue("");
                    } else {
                        cell.setCellValue(value);
                    }
                }

                // 如果类型是BigDecimal
                if ("class java.math.BigDecimal".equals(field.getGenericType().toString())) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    BigDecimal value = (BigDecimal) m.invoke(object);

                    cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                    cell.setCellStyle(cellStyle);
                    if (null == value) {
                        //cell.setCellValue();
                    } else {
                        cell.setCellValue(value.doubleValue());
                    }
                }
                // 如果类型是Long
                if ("class java.lang.Long".equals(field.getGenericType().toString()) || field.getType() == long.class) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    Long value = (Long) m.invoke(object);
                    cell.setCellType(CellType.NUMERIC);
                    if (null == value) {
                        //cell.setCellValue("0");
                    } else {
                        cell.setCellValue(value);
                    }
                }
                // 如果类型是Integer
                if ("class java.lang.Integer".equals(field.getGenericType().toString()) || field.getType() == int.class) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    Integer value = (Integer) m.invoke(object);
                    if (null == value) {
                        //cell.setCellValue("0");
                    } else {
                        cell.setCellType(CellType.NUMERIC);
                        cell.setCellValue(value);
                    }
                }
                // 如果类型是Double
                if ("class java.lang.Double".equals(field.getGenericType().toString()) || field.getType() == double.class) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    Double value = (Double) m.invoke(object);
                    cell.setCellType(CellType.NUMERIC);
                    if (null == value) {
                        //cell.setCellValue("0.00");
                    } else {
                        cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat("0.00"));
                        cell.setCellStyle(cellStyle);
                        cell.setCellValue(value);
                    }
                }
                // 如果类型是Date类型
                if ("class java.util.Date".equals(field.getGenericType().toString())) {
                    Method m = (Method) object.getClass().getMethod(
                            "get" + getMethodName(field.getName()));
                    Date value = (Date) m.invoke(object);
                    if (null == value) {
                        cell.setCellValue("");
                    } else {
                        SimpleDateFormat format = new SimpleDateFormat("yyyy/MM/dd");
                        cell.setCellValue(format.format(value));
                    }
                }
            }
        }
        return wb;
    }


    /**
     * 获取方法名称
     *
     * @param fieldName
     * @return 方法名称
     */
    private static String getMethodName(String fieldName) throws UnsupportedEncodingException {
        byte[] items = fieldName.getBytes("UTF-8");
        if (items[0] >= 97) {
            items[0] = (byte) ((char) items[0] - 'a' + 'A');
        }
        return new String(items, "UTF-8");
    }


    /**
     * 下载Excel文件
     *
     * @param fileName 文件名
     * @param wb       工作簿
     * @param response 输出流
     */
    public static void downExcel(String fileName, Workbook wb, HttpServletResponse response) {
        OutputStream out = null;
        try {
            RequestAttributes requestAttributes = RequestContextHolder.currentRequestAttributes();
            HttpServletRequest request = ((ServletRequestAttributes) requestAttributes).getRequest();
            String userAgent = request.getHeader("User-Agent").toLowerCase();
            response.reset();//清空输出流
            String finalFileName;
            if (org.apache.commons.lang.StringUtils.contains(userAgent, "firefox")) {//火狐浏览器
                fileName = fileName + suffix;
                finalFileName = new String(fileName.getBytes("GB2312"), "ISO8859-1");
                response.setHeader("Content-Disposition", "attachment; filename=\"" + finalFileName + "\"");
                //火狐浏览器设置xlsx格式后缀并弹出下载框
                response.setContentType("application/x-excel");
            } else {
                finalFileName = URLEncoder.encode(fileName, "UTF8");//其他浏览器
                response.setHeader("Content-Disposition", "attachment; filename=" + finalFileName + suffix);
                //设置弹出下载框
                response.setContentType("application/vnd.ms-excel");
            }
            out = response.getOutputStream();
            wb.write(out);
            wb.close();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != out) {
                    out.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }




}
