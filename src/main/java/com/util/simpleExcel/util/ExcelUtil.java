package com.util.simpleExcel.util;

import cn.jointwisdom.mrad.ai.base.annotation.Excel;
import cn.jointwisdom.mrad.ai.base.annotation.ExcelEnum;
import cn.jointwisdom.mrad.ai.base.constant.HabErrorType;
import cn.jointwisdom.mrad.ai.hab.param.ImportParams;
import cn.jointwisdom.mrad.ai.hab.vo.operation.GoalDownloadVO;
import cn.jointwisdom.mrad.ai.hab.vo.operation.OberDownloadVO;
import cn.jointwisdom.mrad.commons.api.constant.RespResult;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.util.Assert;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;

import java.io.*;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.lang.reflect.Type;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 这个类的功能是：
 *
 * @author: zhangyss@foxmail.com
 * @date: 2019/1/25 0025
 */
@Slf4j
public class ExcelUtil {

    private static Logger LG = LoggerFactory.getLogger(ExcelUtil.class);

    public static final String PATTERN = "yyyy/MM/dd";

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式写入系统临时文件
     * 用于多个sheet
     *
     * @param <T>
     * @param sheets {@link ExcelSheet}的集合
     */
    public static <T> File exportExcel(List<ExcelSheet<T>> sheets) {
        return exportExcel(sheets, null);
    }

    public static <T> byte[] exportExcelByte(List<ExcelSheet<T>> sheets) {
        return exportExcelByte(sheets, null);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式写入系统临时文件
     * 用于多个sheet
     *
     * @param sheets      {@link ExcelSheet}的集合
     * @param datePattern 如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    public static <T> File exportExcel(List<ExcelSheet<T>> sheets,
                                       String datePattern) {
        return exportExcel(sheets, null, datePattern);
    }

    public static <T> byte[] exportExcelByte(List<ExcelSheet<T>> sheets,
                                             String datePattern) {
        return exportExcelByte(sheets, null, datePattern);
    }

    /**
     * 利用JAVA的反射机制，将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式写入系统临时文件
     * 用于多个sheet
     *
     * @param <T>
     * @param sheets            {@link ExcelSheet}的集合
     * @param dateTimeFormatter 日期时间格式化
     * @param datePattern       如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     * @return 临时文件
     */
    public static <T> File exportExcel(List<ExcelSheet<T>> sheets, DateTimeFormatter dateTimeFormatter,
                                       String datePattern) {
        Assert.notEmpty(sheets, "sheets不可以为空");
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (ExcelSheet<T> sheet : sheets) {
            // 生成一个表格
            write2Sheet(workbook, sheet, dateTimeFormatter, datePattern);
        }
        OutputStream out = null;
        try {
            String dir = System.getProperty("java.io.tmpdir");
            if (!dir.endsWith(File.separator)) {
                dir = dir + File.separator;
            }
            File file = new File(dir + ExcelUtil.class.getPackage().getName() + System.currentTimeMillis() + ".xlsx");
            out = new FileOutputStream(file);
            workbook.write(out);
            return file;
        } catch (IOException e) {
            LG.error(e.toString(), e);
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    LG.error(e.toString(), e);
                }
            }
        }
        return null;
    }

    public static <T> byte[] exportExcelByte(List<ExcelSheet<T>> sheets, DateTimeFormatter dateTimeFormatter,
                                             String datePattern) {
        Assert.notEmpty(sheets, "sheets不可以为空");
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        for (ExcelSheet<T> sheet : sheets) {
            // 生成一个表格
            write2Sheet(workbook, sheet, dateTimeFormatter, datePattern);
        }
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            workbook.write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            LG.error(e.toString(), e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    LG.error(e.toString(), e);
                }
            }
        }
        return null;
    }


    /**
     * 每个sheet的写入
     *
     * @param workbook          excel对象
     * @param excelSheet        sheet数据集
     * @param dateTimeFormatter 日期时间格式化
     * @param pattern           日期格式
     * @return 临时文件
     */
    private static <T> void write2Sheet(HSSFWorkbook workbook, ExcelSheet<T> excelSheet,
                                        DateTimeFormatter dateTimeFormatter,
                                        String pattern) {
        List<T> dataset = excelSheet.getDataset();
        String sheetName = excelSheet.getSheetName();
        List<String> header = excelSheet.getHeader();
        HSSFSheet sheet = workbook.createSheet(sheetName);

        //时间格式默认"yyyy-MM-dd"
        if (isBlank(pattern)) {
            pattern = PATTERN;
        }
        if (dateTimeFormatter == null) {
            dateTimeFormatter = DateTimeFormatter.ofPattern(pattern);
        }
        //设置样式
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_40_PERCENT.getIndex());
        HSSFFont font = workbook.createFont();
        font.setBold(Boolean.TRUE);
        style.setFont(font);

        if (Boolean.TRUE.equals(excelSheet.oper) && !excelSheet.dataOber.isEmpty()) {
            //绘制标题行
            int rowIndex = 0;   //
            HSSFRow headRow = sheet.createRow(rowIndex);
            Integer headcolumnIndex = 0;
            for (String headName : excelSheet.header) {
                HSSFCell cell = headRow.createCell(headcolumnIndex);
                HSSFRichTextString text = new HSSFRichTextString(headName);
                cell.setCellValue(text);
                cell.setCellStyle(style);
                headcolumnIndex++;
            }
            rowIndex++;

            //绘制内容行
            for (Map<String, List<OberDownloadVO>> stringMapMap : excelSheet.dataOber) {
                for (Map.Entry<String, List<OberDownloadVO>> keyDate : stringMapMap.entrySet()) {
                    int bodyColumnIndex = 0;
                    HSSFRow row = sheet.createRow(rowIndex);
                    HSSFCell cellHead = row.createCell(bodyColumnIndex);
                    setCellValue(cellHead, keyDate.getKey(), dateTimeFormatter, pattern);
                    bodyColumnIndex = 1;
                    for (OberDownloadVO oberDownloadVO : keyDate.getValue()) {
                        Field[] declaredFields = oberDownloadVO.getClass().getDeclaredFields();
                        for (Field field : declaredFields) {
                            Excel annotation = field.getAnnotation(Excel.class);
                            if (annotation != null) {
                                HSSFCell cell = row.createCell(bodyColumnIndex);
                                field.setAccessible(true);
                                Object value = null;
                                try {
                                    value = field.get(oberDownloadVO);
                                } catch (IllegalAccessException e) {
                                    e.printStackTrace();
                                }
                                if (value == null) {
                                    value = "";
                                }
                                setCellValue(cell, value, dateTimeFormatter, pattern);
                                bodyColumnIndex++;
                            }
                        }
                    }
                }
                rowIndex++;
            }

            // 设定自动宽度
            for (int i = 0; i < excelSheet.header.size(); i++) {
                sheet.autoSizeColumn(i);
            }

        } else {
            //无数据创建空白sheet,包含了表头
            if (dataset == null || dataset.isEmpty()) {
                //绘制标题行
                if (!CollectionUtils.isEmpty(header)) {
                    HSSFRow headRow = sheet.createRow(0);
                    Integer headcolumnIndex = 0;
                    for (String s : header) {
                        HSSFCell cell = headRow.createCell(headcolumnIndex);
                        HSSFRichTextString text = new HSSFRichTextString(s);
                        cell.setCellValue(text);
                        cell.setCellStyle(style);
                        // 设定自动宽度
                        sheet.autoSizeColumn(headcolumnIndex);
                        headcolumnIndex++;
                    }
                }
                return;
            }
            //一种是使用map构建，一种是使用bean构建
            //构造序号-列名
            Map<Integer, Field> indexMap = new TreeMap<>();
            Map<Integer, String> indexDefaultValue = new TreeMap<>();
            Map<Integer, String> indexHeaderMap = new TreeMap<>();
            if (excelSheet.getChart()) {
                Field[] declaredFields = dataset.get(0).getClass().getDeclaredFields();
                for (Field field : declaredFields) {
                    Excel annotation = field.getAnnotation(Excel.class);
                    if (annotation != null) {
                        if (indexMap.get(annotation.index()) == null) {
                            indexMap.put(annotation.index(), field);
                            indexDefaultValue.put(annotation.index(), annotation.defaultValue());
                            indexHeaderMap.put(annotation.index(), annotation.name());
                        } else {
                            throw new RuntimeException("重复的列序号!");
                        }
                    }
                }
            } else {
                GoalDownloadVO g = new GoalDownloadVO();
                for (int i = 0; i < header.size(); i++) {
                    for (Field field : g.getClass().getDeclaredFields()) {
                        Excel annotation = field.getAnnotation(Excel.class);
                        if (Objects.nonNull(annotation)) {
                            if (Objects.equals(annotation.name(), header.get(i))) {
                                indexMap.put(i, field);
                                indexDefaultValue.put(i, annotation.defaultValue());
                                indexHeaderMap.put(i, annotation.name());
                            }
                        }
                    }
                }
            }


            //绘制标题行
            int rowIndex = 0;   //
            HSSFRow headRow = sheet.createRow(rowIndex);
            Integer headcolumnIndex = 0;
            for (Map.Entry<Integer, String> entry : indexHeaderMap.entrySet()) {
                HSSFCell cell = headRow.createCell(headcolumnIndex);
                String value = entry.getValue();
                HSSFRichTextString text = new HSSFRichTextString(value);
                cell.setCellValue(text);
                cell.setCellStyle(style);
                headcolumnIndex++;
            }
            rowIndex++;

            //绘制内容行
            for (T t : dataset) {
                int bodyColumnIndex = 0;
                HSSFRow row = sheet.createRow(rowIndex);
                for (Map.Entry<Integer, Field> entry : indexMap.entrySet()) {
                    HSSFCell cell = row.createCell(bodyColumnIndex);
                    Field field = entry.getValue();
                    field.setAccessible(true);
                    Object value = null;
                    try {
                        value = field.get(t);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    if (value == null) {
                        value = indexDefaultValue.get(bodyColumnIndex);
                    }
                    setCellValue(cell, value, dateTimeFormatter, pattern);
                    bodyColumnIndex++;
                }
                rowIndex++;
            }

            if (Objects.nonNull(excelSheet.getTotalAmount())) {
                HSSFRow row = sheet.createRow(rowIndex);
                HSSFCell cell = row.createCell(BigDecimal.ZERO.intValue());
                cell.setCellValue(excelSheet.getTotal());
                cell.setCellStyle(style);
                HSSFCell hssfCell = row.createCell(excelSheet.getIndex());
                hssfCell.setCellValue(excelSheet.getTotalAmount().toString());
                hssfCell.setCellStyle(style);
                //合并列
                CellRangeAddress address = new CellRangeAddress(rowIndex, rowIndex, BigDecimal.ZERO.intValue(),
                        (BigDecimal.valueOf(excelSheet.getIndex()).subtract(BigDecimal.ONE)).intValue());
                sheet.addMergedRegion(address);
                if (excelSheet.getIndex() < (BigDecimal.valueOf(headcolumnIndex).subtract(BigDecimal.ONE)).intValue()) {
                    CellRangeAddress cellAddresses = new CellRangeAddress(rowIndex, rowIndex, excelSheet.getIndex(), (BigDecimal.valueOf(headcolumnIndex).subtract(BigDecimal.ONE)).intValue());
                    sheet.addMergedRegion(cellAddresses);
                }
            }

            // 设定自动宽度
            for (int i = 0; i < indexMap.size(); i++) {
                sheet.autoSizeColumn(i);
            }
        }
    }

    private static void setCellValue(HSSFCell cell, Object value, DateTimeFormatter dateTimeFormatter, String pattern) {
        String textValue = null;
        if (value instanceof Integer) {
            int intValue = (Integer) value;
            cell.setCellValue(intValue);
        } else if (value instanceof Float) {
            float fValue = (Float) value;
            cell.setCellValue(fValue);
        } else if (value instanceof Double) {
            double dValue = (Double) value;
            cell.setCellValue(dValue);
        } else if (value instanceof Long) {
            long longValue = (Long) value;
            cell.setCellValue(longValue);
        } else if (value instanceof Boolean) {
            boolean bValue = (Boolean) value;
            cell.setCellValue(bValue);
        } else if (value instanceof Date) {
            Date date = (Date) value;
            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
            textValue = sdf.format(date);
        } else if (value instanceof LocalDate) {
            if (dateTimeFormatter != null) {
                textValue = ((LocalDate) value).format(dateTimeFormatter);
            } else {
                textValue = value.toString();
            }
        } else if (value instanceof LocalDateTime) {
            if (dateTimeFormatter != null) {
                textValue = ((LocalDateTime) value).format(dateTimeFormatter);
            } else {
                textValue = value.toString();
            }
        } else {
            // 其它数据类型都当作字符串简单处理
            textValue = value == null ? String.valueOf("") : value.toString();
        }
        if (textValue != null) {
            HSSFRichTextString richString = new HSSFRichTextString(textValue);
            cell.setCellValue(richString);
        }
    }

    private static boolean isBlank(String str) {
        if (str == null) {
            return true;
        }
        return str.length() == 0;
    }

    protected static boolean isNotBlank(String str) {
        return !isBlank(str);
    }



    @Data
    public static class ExcelSheet<C> {
        private String sheetName;
        //如果没有数据要设置表头，有数据时，以注解上的name为准
        private List<String> header;
        private List<C> dataset;
        private Boolean oper;
        //订单观察
        private List<Map<String, List<OberDownloadVO>>> dataOber;
        private Integer maxSmartSize;
        //目标达成分析
        private Boolean chart = Boolean.TRUE;
        //账单
        private BigDecimal totalAmount;
        private String total;
        private Integer index;
    }

    /**
     * @Author: chengh@jointwisdom.cn<BR>
     * @Description: Excel导入。对象属性请定义为包装类型 <BR>
     * @Date: 2019/3/14 16:24 <BR>
     * @Param: [is, pojoClass, params] <BR>
     * @return: cn.jointwisdom.mrad.commons.api.constant.RespResult<java.util.List < T>> <BR>
     **/
    public <T> RespResult<List<T>> importExcel(InputStream is, Class<?> pojoClass, ImportParams params) {
        try {
            // 错误信息接收器
            StringBuilder errorMsg = new StringBuilder();
            String fileName = params.getFileName();
            if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
                return HabErrorType.HAB_2412;
            }
            Workbook wb;
            if (fileName.matches("^.+\\.(?i)(xlsx)$")) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = new HSSFWorkbook(is);
            }
            Sheet sheet = wb.getSheetAt(params.getSheetNum());
            if (sheet == null) {
                return HabErrorType.HAB_2414;
            }
            List<T> resultList = new ArrayList<>();
            // 得到Excel的行数
            int totalRows = sheet.getPhysicalNumberOfRows();
            // 总列数
            int totalCells = 0;
            // 得到Excel的列数(前提是有行数)，从第1行算起
            int firstDataRow = params.getHeadRow() + 1;
            if (totalRows >= firstDataRow && sheet.getRow(firstDataRow) != null) {
                totalCells = sheet.getRow(firstDataRow).getPhysicalNumberOfCells();
            } else {
                return HabErrorType.HAB_2413;
            }
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                StringBuilder rowMessage = new StringBuilder();
                Row row = sheet.getRow(r);
                if (row == null) {
                    errorMsg.append("第" + (r + 1) + "行数据有问题,请仔细检查!");
                    continue;
                }
                if (isRowEmpty(row)) {
                    continue;
                }

                T pojo = (T) pojoClass.newInstance();

                // 循环Excel的列
                for (int c = 0; c < totalCells; c++) {
                    Cell cell = row.getCell(c);
                    if (null != cell) {
                        Object cellValue = this.getCellFormatValue(cell);
                        Field[] fields = pojo.getClass().getDeclaredFields();
                        for (Field field : fields) {
                            //若字段值需要转换为其它类型，通过该注解进行转换
                            ExcelEnum annotationEnum = field.getAnnotation(ExcelEnum.class);
                            if (annotationEnum != null) {
                                int index = annotationEnum.index();
                                if (index - 1 == c) {
                                    Type genericType = field.getGenericType();
                                    if ("class java.lang.String".equals(genericType.toString())) {
                                        cellValue = String.valueOf(cellValue);
                                    } else if ("class java.lang.Integer".equals(genericType.toString())) {
                                        cellValue = String.valueOf(cellValue);
                                    } else if ("class java.lang.Boolean".equals(genericType.toString())) {
                                        cellValue = String.valueOf(cellValue);
                                    }

                                    Class aClass = Class.forName(annotationEnum.className());
                                    Method method = aClass.getMethod(annotationEnum.methodName(), String.class);
                                    Object cellEnum = method.invoke(null, cellValue);

                                    field.setAccessible(true);
                                    field.set(pojo, cellEnum);
                                }
                                continue;
                            }

                            Excel annotation = field.getAnnotation(Excel.class);
                            if (annotation != null) {
                                int index = annotation.index();
                                if (index - 1 == c) {
                                    Type genericType = field.getGenericType();
                                    if ("class java.lang.String".equals(genericType.toString())) {
                                        cellValue = String.valueOf(cellValue);
                                    } else if ("class java.time.LocalDate".equals(genericType.toString())) {
                                        cellValue = LocalDateUtil.utilDateToLocalDate((Date) cellValue);
                                    } else if ("class java.time.LocalDateTime".equals(genericType.toString())) {
                                        cellValue = LocalDateUtil.utilDateToLocalDateTime((Date) cellValue);
                                    } else if ("class java.math.BigDecimal".equals(genericType.toString())) {
                                        cellValue = new BigDecimal(String.valueOf(cellValue));
                                    } else if ("class java.util.Date".equals(genericType.toString())) {
                                        cellValue = (Date) cellValue;
                                    } else if ("class java.lang.Integer".equals(genericType.toString())) {
                                        cellValue = Integer.valueOf(String.valueOf(cellValue));
                                    } else if ("class java.lang.Double".equals(genericType.toString())) {
                                        cellValue = Double.valueOf(String.valueOf(cellValue));
                                    } else if ("class java.lang.Boolean".equals(genericType.toString())) {
                                        cellValue = Boolean.valueOf(String.valueOf(cellValue));
                                    } else if ("class java.lang.Short".equals(genericType.toString())) {
                                        cellValue = Short.valueOf(String.valueOf(cellValue));
                                    } else if ("class java.lang.Long".equals(genericType.toString())) {
                                        cellValue = Long.valueOf(String.valueOf(cellValue));
                                    }
                                    field.setAccessible(true);
                                    field.set(pojo, cellValue);
                                }
                            }
                        }
                    }
                }
                // 拼接每行的错误提示
                if (!StringUtils.isEmpty(rowMessage.toString())) {
                    errorMsg.append("第" + (r + 1) + "行," + rowMessage);
                }
                resultList.add(pojo);
            }
            if (!resultList.isEmpty() && StringUtils.isEmpty(errorMsg.toString())) {
                return RespResult.build(resultList);
            }
            return RespResult.build(HabErrorType.HAB_2415, Collections.emptyList(), errorMsg.toString());
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private Object getCellFormatValue(Cell cell) {
        Object cellValue = null;
        if (cell != null) {
            //判断cell类型
            switch (cell.getCellType()) {
                case NUMERIC: {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    } else {
                        //将数值型cell设置为string型
                        cell.setCellType(CellType.STRING);
                        cellValue = cell.getStringCellValue();
                    }
                    break;
                }
                case FORMULA: {
                    //判断cell是否为日期格式
                    if (DateUtil.isCellDateFormatted(cell)) {
                        //转换为日期格式YYYY-mm-dd
                        cellValue = cell.getDateCellValue();
                    } else {
                        //数字
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case STRING: {
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                default:
                    cellValue = "";
            }
        } else {
            cellValue = "";
        }
        return cellValue;
    }

    private boolean isRowEmpty(Row row) {
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                return false;
            }
        }
        return true;
    }

    /**
     * @param excelVo
     * @param sheetName
     * @param dgList
     * @return
     * @description 绘制列数量不固定表格, 该表格需要格式化好内容以list<String>形式传入数据
     * @author yanzhen
     * @date 2020/4/9 17:31
     */
    public static <T> void fillAutoColuSizeDataGrid(ExcelVO excelVo, String sheetName, List<ExcelVO.DataGridVo> dgList, boolean hasBorder) {
        HSSFCellStyle frameStyle = getFrameStyle(excelVo);
        for (int i = 0; i < dgList.size(); i++) {
            ExcelVO.DataGridVo dataGridVo = dgList.get(i);
            //依次绘制每个表格
            List<List<String>> dataList = dataGridVo.getDataList();
            //获取起始行
            int startRowIdx = dataGridVo.getStartRowIdx();
            for (int x = 0; x < dataList.size(); x++) {
                //获取到每一行数据
                List<String> columnList = dataList.get(x);
                //遍历一行数据
                //获取起始列
                int startColIdx = dataGridVo.getStartColIdx();
                for (int index = 0; index < columnList.size(); index++) {
                    HSSFCell cell = excelVo.getCell(sheetName, startRowIdx, startColIdx);
                    //获取当前列内容
                    String cellValue = columnList.get(index);
                    cell.setCellValue(cellValue);
                    if(hasBorder){
                        cell.setCellStyle(frameStyle);
                    }
                    startColIdx++;
                }
                startRowIdx++;
            }
        }
    }

    public static <T> void createSingleRowHeader(ExcelVO excelVo, String sheetName, Class<T> tClass, HSSFCellStyle style) {
        Map<Integer, Field> indexMap = new TreeMap<>();
        Map<Integer, String> indexHeaderMap = new TreeMap<>();
        Field[] declaredFields = tClass.getDeclaredFields();
        for (Field field : declaredFields) {
            if (field.isAnnotationPresent(Excel.class)) {
                Excel annotation = field.getAnnotation(Excel.class);
                if (indexMap.get(annotation.index()) == null) {
                    indexMap.put(annotation.index(), null);
                    indexHeaderMap.put(annotation.index(), annotation.name());
                } else {
                    throw new RuntimeException("重复的列序号!");
                }
            }
        }
        Integer rowIndex = 0;
        Integer headcolumnIndex = 0;
        for (Map.Entry<Integer, String> entry : indexHeaderMap.entrySet()) {
            HSSFCell cell = excelVo.getCell(sheetName, rowIndex, headcolumnIndex);
            String value = entry.getValue();
            HSSFRichTextString text = new HSSFRichTextString(value);
            if (Objects.nonNull(style)) {
                cell.setCellStyle(style);
            }
            cell.setCellValue(text);
            headcolumnIndex++;
        }
        rowIndex++;
    }


    /**
     * @param excelVo
     * @param sheetName
     * @param dgList
     * @return
     * @description 绘制列长度固定的表格，支持多个表格
     * @author yanzhen
     * @date 2020/4/9 14:58
     */
    public static <T> void fillFixColuSizeDataGrid(ExcelVO excelVo, String sheetName, List<ExcelVO.DataGridVo> dgList) {
        //同一个sheet绘制多个表格
        int count = 0;
        for (int i = 0; i < dgList.size(); i++) {
            ExcelVO.DataGridVo dataGridVo = dgList.get(i);
            //顺序
            Map<Integer, Field> indexMap = new TreeMap<>();
            //默认值
            Map<Integer, String> indexDefaultValue = new TreeMap<>();
            Map<Integer, String> indexHeaderMap = new TreeMap<>();
            List<T> dataList = dataGridVo.getDataList();
            /**解析@excel注解 填充内容至对应map中*/
            parseExcelAnno(dataGridVo.getTClass(), indexMap, indexDefaultValue, indexHeaderMap);

            //绘制标题行
            if (dataGridVo.getIncludeBaseheader()) {
                Integer rowIndex = dataGridVo.getStartRowIdx();
                Integer headcolumnIndex = dataGridVo.getStartColIdx();
                for (Map.Entry<Integer, String> entry : indexHeaderMap.entrySet()) {
                    HSSFCell cell = excelVo.getCell(sheetName, rowIndex, headcolumnIndex);
                    String value = entry.getValue();
                    HSSFRichTextString text = new HSSFRichTextString(value);
                    if (Objects.nonNull(dataGridVo.getHeaderStyle())) {
                        cell.setCellStyle(dataGridVo.getHeaderStyle());
                    }
                    cell.setCellValue(text);
                    headcolumnIndex++;
                }
                rowIndex++;
            }
            if (CollectionUtils.isEmpty(dataList)) {
                return;
            }


            int rowIdx = dataGridVo.getIncludeBaseheader() ? dataGridVo.getStartRowIdx() + 1 : dataGridVo.getStartRowIdx();
            HSSFCellStyle frameStyle = getFrameStyle(excelVo);
            for (T t : dataList) {
                int bodyColumnIndex = dataGridVo.getStartColIdx();
                for (Map.Entry<Integer, Field> entry : indexMap.entrySet()) {
                    HSSFCell cell = excelVo.getCell(sheetName, rowIdx, bodyColumnIndex);
                    Field field = entry.getValue();
                    field.setAccessible(true);
                    Object value = null;
                    try {
                        value = field.get(t);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    }
                    if (value == null) {
                        value = indexDefaultValue.get(bodyColumnIndex);
                    }
                    if(dataGridVo.getHasBorder()){
                        cell.setCellStyle(frameStyle);
                        count++;
                    }
                    setCellValue(cell, value, dataGridVo.getDateTimeFormatter(), dataGridVo.getDatePattern());
                    bodyColumnIndex++;
                }
                rowIdx++;

            }
            System.out.println("总共设置了:"+count);
            log.info("总共设置了:"+count);
        }


    }

    /**
     * @param tClass
     * @param indexMap
     * @param indexDefaultValue
     * @param indexHeaderMap
     * @return
     * @description 解析输入对象的@Excel注解 输出对应map
     * @author yanzhen
     * @date 2020/4/9 14:06
     */
    private static <T> void parseExcelAnno(Class<T> tClass, Map<Integer, Field> indexMap, Map<Integer, String> indexDefaultValue, Map<Integer, String> indexHeaderMap) {
        Field[] declaredFields = tClass.getDeclaredFields();
        for (Field field : declaredFields) {
            Excel annotation = field.getAnnotation(Excel.class);
            if (annotation != null) {
                if (indexMap.get(annotation.index()) == null) {
                    indexMap.put(annotation.index(), field);
                    indexDefaultValue.put(annotation.index(), annotation.defaultValue());
                    indexHeaderMap.put(annotation.index(), annotation.name());
                } else {
                    throw new RuntimeException("重复的列序号!");
                }
            }
        }
    }


    /**
     * @param excelVo
     * @param sheetName
     * @param cellRangeVoList
     * @return
     * @description 单元格合并功能 用于绘制表头表尾
     * @author yanzhen
     * @date 2020/4/9 13:20
     */
    public static void mergeCell(ExcelVO excelVo, String sheetName, List<ExcelVO.CellRangeVo> cellRangeVoList) {
        HSSFSheet sheet = excelVo.getSheet(sheetName);
        for (int i = 0; i < cellRangeVoList.size(); i++) {
            ExcelVO.CellRangeVo cellRangeVo = cellRangeVoList.get(i);
            HSSFCell cell = excelVo.getCell(sheetName, cellRangeVo.getFirstRow(), cellRangeVo.getFirstCol());
            if (Objects.nonNull(cellRangeVo.getStyle())) {
                cell.setCellStyle(cellRangeVo.getStyle());
            }
            cell.setCellValue(cellRangeVo.getCellText());
            CellRangeAddress address = new CellRangeAddress(cellRangeVo.getFirstRow(), cellRangeVo.getLastRow(),
                    cellRangeVo.getFirstCol(), cellRangeVo.getLastCol());
            sheet.addMergedRegion(address);
        }
    }

    /**
     * @param excelVo
     * @param sheetName
     * @param startColumnIdx
     * @param endColumnIdx
     * @return
     * @description 设置列的自适应
     * @author yanzhen
     * @date 2020/4/9 16:29
     */
    public static void setColumnAutoSize(ExcelVO excelVo, String sheetName, int startColumnIdx, int endColumnIdx) {
        HSSFSheet sheet = excelVo.getSheet(sheetName);
        for (int i = startColumnIdx; i <= endColumnIdx; i++) {
            sheet.autoSizeColumn(i);
        }
    }


    /**
     * @param excelVo
     * @param sheetName
     * @param list
     * @return
     * @description 设置单元格格式
     * @author yanzhen
     * @date 2020/4/9 16:29
     */
    public static void changeCellStyle(ExcelVO excelVo, String sheetName, List<ExcelVO.CellStyleVo> list) {
        for (int i = 0; i < list.size(); i++) {
            ExcelVO.CellStyleVo cellStyleVo = list.get(i);
            HSSFCell cell = excelVo.getCell(sheetName, cellStyleVo.getRowIdx(), cellStyleVo.getColIdx());
            cell.setCellStyle(cellStyleVo.getStyle());
        }
    }


    /**
     * @param excel
     * @return
     * @description excel工作簿转换File对象
     * @author yanzhen
     * @date 2020/4/9 19:02
     */
    public static File excel2File(ExcelVO excel) {
        OutputStream out = null;
        try {
            String dir = System.getProperty("java.io.tmpdir");
            if (!dir.endsWith(File.separator)) {
                dir = dir + File.separator;
            }
            File file = new File(dir + ExcelUtil.class.getPackage().getName() + System.currentTimeMillis() + ".xlsx");
            out = new FileOutputStream(file);
            excel.getWorkbook().write(out);
            return file;
        } catch (IOException e) {
            LG.error(e.toString(), e);
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    LG.error(e.toString(), e);
                }
            }
        }
        return null;
    }

    /**
     * @param excel
     * @return
     * @description excel工作簿转换byte数组
     * @author yanzhen
     * @date 2020/4/9 19:02
     */
    public static byte[] excel2Byte(ExcelVO excel) {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            excel.getWorkbook().write(outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            LG.error(e.toString(), e);
        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    LG.error(e.toString(), e);
                }
            }
        }
        return null;
    }

    /**
     * @param excelVO
     * @return
     * @description 获取表头默认样式
     * @author yanzhen
     * @date 2020/4/10 11:40
     */
    public static HSSFCellStyle getDefaultHeaderStyle(ExcelVO excelVO,boolean hasBorder) {
        HSSFCellStyle style = excelVO.getWorkbook().createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);// 左右居中
        style.setVerticalAlignment(VerticalAlignment.CENTER);// 上下居中
        if(hasBorder){
            style.setBorderBottom(BorderStyle.THIN); //下边框
            style.setBorderLeft(BorderStyle.THIN);//左边框
            style.setBorderTop(BorderStyle.THIN);//上边框
            style.setBorderRight(BorderStyle.THIN);//右边框
        }
        HSSFFont font = excelVO.getWorkbook().createFont();
        font.setBold(Boolean.TRUE);
        style.setFont(font);
        return style;
    }

    /**
     * @param excelVO
     * @return
     * @description 获取默认灰色单元格背景
     * @author yanzhen
     * @date 2020/4/10 15:24
     */
    public static HSSFCellStyle getGreyColorCell(ExcelVO excelVO,boolean hasBorder) {
        HSSFCellStyle style = excelVO.getWorkbook().createCellStyle();
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        if(hasBorder){
            style.setBorderBottom(BorderStyle.THIN); //下边框
            style.setBorderLeft(BorderStyle.THIN);//左边框
            style.setBorderTop(BorderStyle.THIN);//上边框
            style.setBorderRight(BorderStyle.THIN);//右边框
        }
        return style;
    }

    /**
     * @description 设置样式
     * @author yanzhen
     * @date 2020/4/10 15:56
     * @param excelVO
     * @param sheetName
     * @param startRowIdx
     * @param endRowIdx
     * @param startColIdx
     * @param endColIdx
     * @return
     */
    public static void changeDgBorderStyle(ExcelVO excelVO, String sheetName, int startRowIdx, int endRowIdx, int startColIdx, int endColIdx,HSSFCellStyle cellStyle) {
        for (int row = startRowIdx; row <= endRowIdx; row++) {
            for (int col = startColIdx ;col<=endColIdx;col++){
                HSSFCell cell = excelVO.getCell(sheetName, row, col);
                cell.setCellStyle(cellStyle);
            }
        }
    }


    /**
     * @description 获取带边框普通单元格样式
     * @author yanzhen
     * @date 2020/4/10 16:13
     * @param
     * @return
     */
    public static HSSFCellStyle getFrameStyle(ExcelVO excelVO) {
        HSSFCellStyle cellStyle = excelVO.getWorkbook().createCellStyle();
        cellStyle.setBorderBottom(BorderStyle.THIN); //下边框
        cellStyle.setBorderLeft(BorderStyle.THIN);//左边框
        cellStyle.setBorderTop(BorderStyle.THIN);//上边框
        cellStyle.setBorderRight(BorderStyle.THIN);//右边框
        return  cellStyle;
    }

}
