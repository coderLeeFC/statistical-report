package cn.sljl.util;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.util.List;

/**
 * 标题模型
 *
 * @author wangeqiu
 * @version 1.0.0
 * @date 2024/04/09 09:22:52
 */
public class TitleModel {

    /**
     * 创建营业收入&营业总成本&毛利率（sheet）标题
     *
     * @param workbook excel workbook
     * @param sheet    excel sheet
     * @param size     excel宽度
     * @author wangeqiu
     * @date 2024/04/09 09:22:52
     */
    public void createMBTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size){
        /**
         * 第2行
         */
        //设置“填报单位（公章）”、“单位：万元”单元格合并
        createSecondTitle(workbook,sheet,size);

        /**
         * 第3-6行左侧样式
         */
        XSSFRow thirdRow = sheet.createRow(2);
        createThirdTitle(workbook, sheet,thirdRow);

        //营业收入
        CellRangeAddress cellAddress33 = new CellRangeAddress(2, 2, 2, 13);
        sheet.addMergedRegion(cellAddress33);
        XSSFCell cell33 = thirdRow.createCell(2);
        cell33.setCellStyle(cellCommonStyle(workbook));
        cell33.setCellValue(Constant.TITLE_SECOND_RIGHT.get(0));

        CellRangeAddress cellAddress41 = new CellRangeAddress(3, 3, 2, 7);
        CellRangeAddress cellAddress42 = new CellRangeAddress(3, 3, 8, 13);

        sheet.addMergedRegion(cellAddress41);
        sheet.addMergedRegion(cellAddress42);

        XSSFRow forthRow = sheet.createRow(3);

        XSSFCell cell41 = forthRow.createCell(2);
        XSSFCell cell42 = forthRow.createCell(8);

        cell41.setCellStyle(cellCommonStyle(workbook));
        cell41.setCellValue(Constant.TITLE_THIRD_RIGHT_MB.get(0));
        cell42.setCellStyle(cellCommonStyle(workbook));
        cell42.setCellValue(Constant.TITLE_THIRD_RIGHT_MB.get(1));

        //营业总成本
        CellRangeAddress cellAddress34 = new CellRangeAddress(2, 3, 14, 19);
        sheet.addMergedRegion(cellAddress34);

        XSSFCell cell34 = thirdRow.createCell(14);
        cell34.setCellStyle(cellCommonStyle(workbook));
        cell34.setCellValue(Constant.TITLE_SECOND_RIGHT.get(1));

        //毛利率
        CellRangeAddress cellAddress35 = new CellRangeAddress(2, 3, 20, size-1);
        sheet.addMergedRegion(cellAddress35);

        XSSFCell cell35 = thirdRow.createCell(20);
        cell35.setCellStyle(cellCommonStyle(workbook));
        cell35.setCellValue(Constant.TITLE_THIRD_RIGHT_MB.get(3));

        /**
         * 第5行右侧样式
         */
        createFifthTitle(workbook, sheet, size);

        /**
         * 第6行右侧样式
         */
        createSixthTitle(workbook, sheet, size);

    }

    /**
     * 创建现金流（sheet）标题
     *
     * @param workbook excel workbook
     * @param sheet    excel sheet
     * @param size     excel宽度
     * @param title    标题
     * @param title1   标题1
     * @author wangeqiu
     * @date 2024/04/09 09:22:53
     */
    public void createCFSTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size,String title,List<String> title1){
        /**
         * 第2行
         */
        //设置“填报单位（公章）”、“单位：万元”单元格合并
        createSecondTitle(workbook,sheet,size);

        /**
         * 第3-6行左侧样式
         */
        XSSFRow thirdRow = sheet.createRow(2);
        createThirdTitle(workbook, sheet,thirdRow);

        //现金流
        CellRangeAddress cellAddress33 = new CellRangeAddress(2, 2, 2, size-1);
        sheet.addMergedRegion(cellAddress33);
        XSSFCell cell33 = thirdRow.createCell(2);
        cell33.setCellStyle(cellCommonStyle(workbook));
        cell33.setCellValue(title);

        XSSFRow forthRow = sheet.createRow(3);

        for (int i = 0; i < (size-2)/6; i++) {
            CellRangeAddress cellAddress = new CellRangeAddress(3, 3, 2+6*i, 7+6*i);
            sheet.addMergedRegion(cellAddress);

            XSSFCell cell = forthRow.createCell(2+6*i);
            cell.setCellStyle(cellCommonStyle(workbook));
            cell.setCellValue(title1.get(i));

        }

        /**
         * 第5行右侧样式
         */
        createFifthTitle(workbook, sheet, size);

        /**
         * 第6行右侧样式
         */
        createSixthTitle(workbook, sheet, size);
    }

    /**
     * 创建工资性支出（sheet）标题
     *
     * @param workbook excel workbook
     * @param sheet    excel sheet
     * @param size     excel宽度
     * @param title    标题
     * @author wangeqiu
     * @date 2024/04/09 09:22:53
     */
    public void createSalaryTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size,String title){
        /**
         * 第2行
         */
        //设置“填报单位（公章）”、“单位：万元”单元格合并
        createSecondTitle(workbook,sheet,size);

        /**
         * 第3-6行左侧样式
         */
        XSSFRow thirdRow = sheet.createRow(2);
        CellRangeAddress cellAddress31 =  new CellRangeAddress(2, 4, 0, 0);
        CellRangeAddress cellAddress32 = new CellRangeAddress(2, 4, 1, 1);

        sheet.addMergedRegion(cellAddress31);
        sheet.addMergedRegion(cellAddress32);

        XSSFCell cell31 = thirdRow.createCell(0);
        XSSFCell cell32 = thirdRow.createCell(1);

        cell31.setCellStyle(cellCommonStyle(workbook));
        cell31.setCellValue(Constant.TITLE_SECOND.get(0));
        cell32.setCellStyle(cellCommonStyle(workbook));
        cell32.setCellValue(Constant.TITLE_SECOND.get(1));

        CellRangeAddress cellAddress33 = new CellRangeAddress(2, 2, 2, size-1);
        sheet.addMergedRegion(cellAddress33);
        XSSFCell cell33 = thirdRow.createCell(2);
        cell33.setCellStyle(cellCommonStyle(workbook));
        cell33.setCellValue(title);


        /**
         * 第4行右侧样式
         */
        XSSFRow fifthRow = sheet.createRow(3);
        for (int i = 0; i < (size-2)/2; i++) {
            CellRangeAddress cellAddress53 = new CellRangeAddress(3, 3, 2 + 2*i, 3 + 2*i);

            sheet.addMergedRegion(cellAddress53);

            XSSFCell cell53 = fifthRow.createCell(2 + 2*i);

            if (i<3){
                cell53.setCellStyle(cellCommonStyle(workbook));
                cell53.setCellValue(Constant.TITLE_SEVENTH.get(i));
            }else {
                cell53.setCellStyle(cellCommonStyle(workbook));
                cell53.setCellValue(Constant.TITLE_SEVENTH.get(i%3));
            }
        }

        /**
         * 第5行右侧样式
         */
        XSSFRow sixthRow = sheet.createRow(4);
        for (int i = 2; i < size; i++) {
            XSSFCell cell = sixthRow.createCell(i);

            if (i%2==0){
                cell.setCellStyle(cellCommonStyle(workbook));
                cell.setCellValue(Constant.TITLE_EIGHTH.get(0));
            }else {
                cell.setCellStyle(cellCommonStyle(workbook));
                cell.setCellValue(Constant.TITLE_EIGHTH.get(1));
            }
        }
    }

    /**
     * 创建第二行通用格式【填报单位（公章）、单位：万元】
     *
     * @param workbook excel workbook
     * @param sheet    excel sheet
     * @param size     excel宽度
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public void createSecondTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size){
        //设置“填报单位（公章）”、“单位：万元”单元格合并
        CellRangeAddress cellAddress21 = new CellRangeAddress(1, 1, 0, 2);
        CellRangeAddress cellAddress22 = new CellRangeAddress(1, 1, size-2, size-1);

        sheet.addMergedRegion(cellAddress21);
        sheet.addMergedRegion(cellAddress22);

        //设置值
        XSSFRow secondRow = sheet.createRow(1);

        XSSFCell cell21 = secondRow.createCell(0);
        XSSFCell cell22 = secondRow.createCell(size-2);

        cell21.setCellStyle(cellCommonStyle(workbook));
        cell21.setCellValue(Constant.TITLE_FIRST.get(0));
        cell22.setCellStyle(cellCommonStyle(workbook));
        cell22.setCellValue(Constant.TITLE_FIRST.get(1));
    }

    /**
     * 创建通用格式【序号、业务板块】
     *
     * @param workbook excel workbook
     * @param sheet    excel sheet
     * @param thirdRow excel row
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public void createThirdTitle(XSSFWorkbook workbook, XSSFSheet sheet,XSSFRow thirdRow){
        CellRangeAddress cellAddress31 =  new CellRangeAddress(2, 5, 0, 0);
        CellRangeAddress cellAddress32 = new CellRangeAddress(2, 5, 1, 1);

        sheet.addMergedRegion(cellAddress31);
        sheet.addMergedRegion(cellAddress32);

        XSSFCell cell31 = thirdRow.createCell(0);
        XSSFCell cell32 = thirdRow.createCell(1);

        cell31.setCellStyle(cellCommonStyle(workbook));
        cell31.setCellValue(Constant.TITLE_SECOND.get(0));
        cell32.setCellStyle(cellCommonStyle(workbook));
        cell32.setCellValue(Constant.TITLE_SECOND.get(1));
    }

    /**
     * 创建第五个标题
     *
     * @param workbook 工作簿
     * @param sheet    床单
     * @param size     大小
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public void createFifthTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size){
        XSSFRow fifthRow = sheet.createRow(4);
        for (int i = 0; i < (size-2)/2; i++) {
            CellRangeAddress cellAddress53 = new CellRangeAddress(4, 4, 2 + 2*i, 3 + 2*i);

            sheet.addMergedRegion(cellAddress53);

            XSSFCell cell53 = fifthRow.createCell(2 + 2*i);

            if (i<3){
                cell53.setCellStyle(cellCommonStyle(workbook));
                cell53.setCellValue(Constant.TITLE_SEVENTH.get(i));
            }else {
                cell53.setCellStyle(cellCommonStyle(workbook));
                cell53.setCellValue(Constant.TITLE_SEVENTH.get(i%3));
            }
        }
    }

    /**
     * 创建第六个标题
     *
     * @param workbook 工作簿
     * @param sheet    床单
     * @param size     大小
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public void createSixthTitle(XSSFWorkbook workbook, XSSFSheet sheet, int size){
        XSSFRow sixthRow = sheet.createRow(5);
        for (int i = 2; i < size; i++) {
            XSSFCell cell = sixthRow.createCell(i);

            if (i%2==0){
                cell.setCellStyle(cellCommonStyle(workbook));
                cell.setCellValue(Constant.TITLE_EIGHTH.get(0));
            }else {
                cell.setCellStyle(cellCommonStyle(workbook));
                cell.setCellValue(Constant.TITLE_EIGHTH.get(1));
            }
        }
    }

    /**
     * 设置字体为宋体，大小为14，加粗
     * 设置单元格垂直、水平居中
     *
     * @param workbook
     * @return {@link XSSFCellStyle }
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public XSSFCellStyle cellCommonStyle(XSSFWorkbook workbook){
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));

        return cellStyle;
    }

    /**
     * 设置字体为宋体，大小为14
     * 设置单元格垂直、水平居中
     *
     * @param workbook
     * @return {@link XSSFCellStyle }
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public XSSFCellStyle cellCommonStyle1(XSSFWorkbook workbook){
        XSSFFont font = workbook.createFont();
//        font.setBold(true);
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));

        return cellStyle;
    }

    /**
     * 设置字体为宋体，大小为14，加粗
     * 设置单元格垂直居中、水平居右
     *
     * @param workbook
     * @return {@link XSSFCellStyle }
     * @author wangeqiu
     * @date 2024/04/09 09:22:54
     */
    public XSSFCellStyle cellCommonStyle2(XSSFWorkbook workbook){
        XSSFFont font = workbook.createFont();
        font.setBold(true);
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));

        return cellStyle;
    }

    /**
     * 设置字体为宋体，大小为14
     * 设置单元格垂直居中、水平居右
     *
     * @param workbook
     * @return {@link XSSFCellStyle }
     * @author wangeqiu
     * @date 2024/04/09 09:22:55
     */
    public XSSFCellStyle cellCommonStyle3(XSSFWorkbook workbook){
        XSSFFont font = workbook.createFont();
//        font.setBold(true);
        font.setFontName("宋体");
        font.setFontHeightInPoints((short) 14);

        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.RIGHT);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setDataFormat(workbook.createDataFormat().getFormat("0"));

        return cellStyle;
    }

}
