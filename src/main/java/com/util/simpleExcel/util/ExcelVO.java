package com.util.simpleExcel.util;

import com.google.common.collect.Maps;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.springframework.util.Assert;

import java.time.format.DateTimeFormatter;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * @className: ExcelVo
 * @description: excel构建辅助VO, 每个vo对应一个excel
 * @author: YanZhen
 * @date: 2020/4/9 10:33
 * @version: 1.0
 */
public class ExcelVO {

    private final HSSFWorkbook workbook = new HSSFWorkbook();


    /**Sheet存储容器*/
    private Map<String, HSSFSheet> sheetsContainer = Maps.newLinkedHashMap();
    /**Row存储容器*/
    private Map<String, HSSFRow> rowContainer = Maps.newLinkedHashMap();
    /**Cell存储容器*/
    private Map<String, Map<Integer, HSSFCell>> cellContainer = Maps.newLinkedHashMap();

    /**
     * @description 获取当前excel工作簿对象
     * @author yanzhen
     * @date 2020/4/9 12:56
     * @param
     * @return
     */
    public HSSFWorkbook getWorkbook(){
        return workbook;
    }

    /**
     * @param sheetName
     * @return
     * @description 根据名称获取sheet对象,没有则创建
     * @author yanzhen
     * @date 2020/4/9 10:42
     */
    public HSSFSheet getSheet(String sheetName) {
        Assert.notNull(sheetName, "sheet名称不能为空");
        if (sheetsContainer.containsKey(sheetName)) {
            return sheetsContainer.get(sheetName);
        }
        HSSFSheet sheet = workbook.createSheet(sheetName);
        sheetsContainer.put(sheetName,sheet);
        return  sheet;
    }

    /**
     * @description 根据名称获取Row对象,没有则创建
     * @author yanzhen
     * @date 2020/4/9 11:28
     * @param sheetName
     * @param rowIndex
     * @return
     */
    public HSSFRow getRow(String sheetName,Integer rowIndex){
        Assert.state(rowIndex >= 0,"行数不能为负值");
        Assert.notNull(sheetName, "sheet名称不能为空");
        String rowContainerKey = getSnameRindexKey(sheetName, rowIndex);
        if(rowContainer.containsKey(rowContainerKey)){
            return rowContainer.get(rowContainerKey);
        }
        HSSFSheet sheet = getSheet(sheetName);
        HSSFRow row = sheet.createRow(rowIndex);
        rowContainer.put(rowContainerKey,row);
        return row;
    }

    /**
     * @description 创建单元格
     * @author yanzhen
     * @date 2020/4/9 12:50
     * @param sheetName
     * @param rowIndex
     * @param cellColumn
     * @return
     */
    public HSSFCell createCell(String sheetName,Integer rowIndex,Integer cellColumn){
        Assert.state(rowIndex >= 0,"行数不能为负值");
        Assert.state(cellColumn >= 0,"列数不能为负值");
        Assert.notNull(sheetName, "sheet名称不能为空");
        String cellContainerKey = getSnameRindexKey(sheetName, rowIndex);
        if(!sheetsContainer.containsKey(sheetName)){
           getSheet(sheetName);
        }
        if(!rowContainer.containsKey(rowIndex)){
           getRow(sheetName,rowIndex);
        }
        if(cellContainer.containsKey(cellContainerKey)){
            Map<Integer, HSSFCell> columnHSSFCellMap = cellContainer.get(cellContainerKey);
            if(columnHSSFCellMap.containsKey(cellColumn)){
                return columnHSSFCellMap.get(cellColumn);
            }else{
                HSSFRow row = rowContainer.get(cellContainerKey);
                HSSFCell cell = row.createCell(cellColumn);
                columnHSSFCellMap.put(cellColumn,cell);
                return cell;
            }
        }else{
            Map<Integer, HSSFCell> columnHSSFCellMap =  Maps.newLinkedHashMap();
            HSSFRow row = rowContainer.get(cellContainerKey);
            HSSFCell cell = row.createCell(cellColumn);
            columnHSSFCellMap.put(cellColumn,cell);
            cellContainer.put(cellContainerKey,columnHSSFCellMap);
            return cell;
        }
    }

    /**
     * @description 获取单元格对象，没有则创建
     * @author yanzhen
     * @date 2020/4/9 12:57
     * @param sheetName
     * @param rowIndex
     * @param cellColumn
     * @return
     */
    public HSSFCell getCell(String sheetName,Integer rowIndex,Integer cellColumn){
        Assert.state(rowIndex >= 0,"行数不能为负值");
        Assert.state(cellColumn >= 0,"列数不能为负值");
        Assert.notNull(sheetName, "sheet名称不能为空");
        String cellContainerKey = getSnameRindexKey(sheetName, rowIndex);
        if(cellContainer.containsKey(cellContainerKey)){
            Map<Integer, HSSFCell> coluHSSFCellMap = cellContainer.get(cellContainerKey);
            if(coluHSSFCellMap.containsKey(cellColumn)){
                return coluHSSFCellMap.get(cellColumn);
            }else{
                return createCell(sheetName,rowIndex,cellColumn);
            }
        }else{
            return createCell(sheetName,rowIndex,cellColumn);
        }
    }

    /**生成rowContainer和cellContainer的key*/
    private String getSnameRindexKey(String sheetName, Integer rowIndex) {
        return StringUtil.join("-", Arrays.asList(sheetName, rowIndex));
    }


    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class CellRangeVo{
        private int firstRow;
        private int lastRow;
        private int firstCol;
        private int lastCol;
        private String cellText;
        private HSSFCellStyle style;

    }

    @Data
    @AllArgsConstructor
    @NoArgsConstructor
    public static class CellStyleVo{
        private int rowIdx;
        private int colIdx;
        private HSSFCellStyle style;
    }

    @Data
    public static class DataGridVo<C>{
        private int startRowIdx;
        private int startColIdx;
        private DateTimeFormatter dateTimeFormatter;
        private String datePattern;
        // true标识绘制默认表头
        private Boolean includeBaseheader = Boolean.FALSE;
        // true 表示绘制单元格带边框
        private Boolean hasBorder = Boolean.FALSE;
        // 绘制默认表头时的样式 为null时不单独设置样式
        private HSSFCellStyle headerStyle;
        private Class<C> tClass;
        private List<C> dataList;

    }

}
