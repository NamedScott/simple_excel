package com.util.simpleExcel;

import com.util.simpleExcel.util.ExcelUtil;
import com.util.simpleExcel.util.ExcelVO;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

/**
 * @className: ExcelTest
 * @description: //TODO
 * @author: YanZhen
 * @date: 2020/4/14 14:23
 * @version: 1.0
 */
@RunWith(SpringRunner.class)
@SpringBootTest
@Slf4j
public class ExcelTest {

    @Test
    public void createChangeableColumAmount() throws IOException {
        /**绘制列数量动态变化表格*/
        ExcelVO excelVo = new ExcelVO();
        String sheetName = "第一个sheet";

        //sheet1绘制表头
        ExcelVO.CellRangeVo cellRangeVo1 = new ExcelVO.CellRangeVo(0, 0, 0, 3, "酒店名称:北京青蓝大酒店",null);
        ExcelVO.CellRangeVo cellRangeVo2 = new ExcelVO.CellRangeVo(1, 1, 0, 3, "发起日期:2019/12/01-2019/12/30",null);
        ExcelUtil.mergeCell(excelVo, sheetName, Arrays.asList(cellRangeVo1,cellRangeVo2));

        //使用动态方法绘制表格
        List<String> list = Arrays.asList("1","2","3","4","5","6","7","8","9");
        List<String> list1 = Arrays.asList("11","22","33","44","55","66","77","88","99");
        List<String> list2 = Arrays.asList("单元格1","单元格2","单元格3","单元格4","单元格5","单元格6","单元格7","单元格8","单元格9");
        List<String> list3 = Arrays.asList("单元格1","单元格2","单元格3","单元格4","单元格5","单元格6","单元格7","单元格8");
        ExcelVO.DataGridVo dataGridVo = new ExcelVO.DataGridVo();
        dataGridVo.setStartRowIdx(5);
        dataGridVo.setStartColIdx(0);
        dataGridVo.setDataList(Arrays.asList(list2,list3,list,list1));
        ExcelUtil.fillAutoColuSizeDataGrid(excelVo,sheetName,Arrays.asList(dataGridVo),true);
        HSSFCellStyle frameStyle = ExcelUtil.getFrameStyle(excelVo);
        //ExcelUtil.changeDgBorderStyle(excelVo,sheetName,6, dataGridVo.getDataList().size()+6,0,list.size()+1,frameStyle);


        //设置单元格自适应
        ExcelUtil.setColumnAutoSize(excelVo,sheetName,0,20);

        HSSFWorkbook workbook = excelVo.getWorkbook();
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);

        File file = new File("D:\\"+"非固定列数量"+".xls");
        try {
            file.createNewFile();
        } catch (IOException e) {
            e.printStackTrace();
        }
        try(
                FileOutputStream fileOutputStream = new FileOutputStream(file);) {

            fileOutputStream.write(outputStream.toByteArray());
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
        }
    }

}
