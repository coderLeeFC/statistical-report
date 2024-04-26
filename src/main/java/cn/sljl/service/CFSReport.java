package cn.sljl.service;

import cn.sljl.mapper.CommonSql;
import cn.sljl.util.Constant;
import cn.sljl.util.DateUtils;
import cn.sljl.util.DeptUtils;
import cn.sljl.util.TitleModel;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.Connection;
import java.sql.SQLException;

/**
 * 现金流
 *
 * @author wangeqiu
 * @version 1.0.0
 * @date 2024/04/09 13:59:55
 */
public class CFSReport {
    /**
     * 主方法
     *
     * @param connection 数据库链接
     * @param workbook   excel workbook
     * @param sheet      excel sheet
     * @param titleModel 标题
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:55
     */
    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        DateUtils dateUtils = new DateUtils();
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();

        titleModel.createCFSTitle(workbook, sheet, 20,Constant.CFS_TITLE_RIGHT,Constant.TITLE_SIXTH_RIGHT_CFS);

        getSljlOCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getSljlMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZbdlMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZjzxMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHhakCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHhchMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHyjcCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getSdsjCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getPgzxCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash1(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash2(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash3(connection,workbook,sheet,dateUtils,commonSql, deptUtils);

    }

    /**
     * 缴纳税费【只有流出】
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:55
     */
    private void getOtherCash3(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //监理
        //企业所得税
        double sljlTytm1=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlTy1  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlLytm1=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlLy1  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        //个人所得税
        double sljlTytm2=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlTy2  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlLytm2=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlLy2  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        //房产税
        double sljlTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        //城镇土地使用税
        double sljlTytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlTy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlLytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlLy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        //印花税
        double sljlTytm5=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlTy5  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlLytm5=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlLy5  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");

        //华海
        double hhakTytm1= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakTy1  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakLytm1= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakLy1  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakTytm2= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakTy2  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakLytm2= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakLy2  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakTytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakTy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakLytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakLy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakTytm5= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakTy5  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakLytm5= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakLy5  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222108");

        //恒远
        double hyjcTytm1= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcTy1  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcLytm1= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcLy1  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcTytm2= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcTy2  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcLytm2= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcLy2  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcTytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcTy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcLytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcLy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcTytm5= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcTy5  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcLytm5= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcLy5  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222108");

        //设计
        double sdsjTytm1=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjTy1  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjLytm1=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjLy1  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjTytm2=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjTy2  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjLytm2=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjLy2  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjTytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjTy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjLytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjLy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjTytm5=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjTy5  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjLytm5=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjLy5  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");

        double tytmOut=sljlTytm1+sljlTytm2+sljlTytm3+sljlTytm4+sljlTytm5+hhakTytm1+hhakTytm2+hhakTytm3+hhakTytm4+hhakTytm5+hyjcTytm1+hyjcTytm2+hyjcTytm3+hyjcTytm4+hyjcTytm5+sdsjTytm1+sdsjTytm2+sdsjTytm3+sdsjTytm4+sdsjTytm5;
        double tyOut  =sljlTy1  +sljlTy2  +sljlTy3  +sljlTy4  +sljlTy5  +hhakTy1  +hhakTy2  +hhakTy3  +hhakTy4  +hhakTy5  +hyjcTy1  +hyjcTy2  +hyjcTy3  +hyjcTy4  +hyjcTy5  +sdsjTy1  +sdsjTy2  +sdsjTy3  +sdsjTy4  +sdsjTy5  ;
        double lytmOut=sljlLytm1+sljlLytm2+sljlLytm3+sljlLytm4+sljlLytm5+hhakLytm1+hhakLytm2+hhakLytm3+hhakLytm4+hhakLytm5+hyjcLytm1+hyjcLytm2+hyjcLytm3+hyjcLytm4+hyjcLytm5+sdsjLytm1+sdsjLytm2+sdsjLytm3+sdsjLytm4+sdsjLytm5;
        double lyOut  =sljlLy1  +sljlLy2  +sljlLy3  +sljlLy4  +sljlLy5  +hhakLy1  +hhakLy2  +hhakLy3  +hhakLy4  +hhakLy5  +hyjcLy1  +hyjcLy2  +hyjcLy3  +hyjcLy4  +hyjcLy5  +sdsjLy1  +sdsjLy2  +sdsjLy3  +sdsjLy4  +sdsjLy5  ;

        deptUtils.outputSome(Constant.TAX,workbook,sheet,34,
                tytmOut,
                tyOut  ,
                lytmOut,
                lyOut  );
    }

    /**
     * 备用金借款
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:55
     */
    private void getOtherCash2(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入【其他应收款-个人（贷方）】
        //监理
        double sljlTytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlTy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        //华海
        double hhakTytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakTy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        //恒远
        double hyjcTytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcTy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        //设计
        double sdsjTytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjTy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");

        double tytmIn = sljlTytm1+hhakTytm1+sdsjTytm1+hyjcTytm1;
        double tyIn   = sljlTy1  +hhakTy1  +sdsjTy1  +hyjcTy1  ;
        double lytmIn = sljlLytm1+hhakLytm1+sdsjLytm1+hyjcLytm1;
        double lyIn   = sljlLy1  +hhakLy1  +sdsjLy1  +hyjcLy1  ;

        //现金流出【其他应收款-个人（借方）】
        //监理
        double sljlTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        //华海
        double hhakTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        //恒远
        double hyjcTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        //设计
        double sdsjTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");

        double tytmOut=sljlTytm3+hhakTytm3+hyjcTytm3+sdsjTytm3;
        double tyOut  =sljlTy3  +hhakTy3  +hyjcTy3  +sdsjTy3  ;
        double lytmOut=sljlLytm3+hhakLytm3+hyjcLytm3+sdsjLytm3;
        double lyOut  =sljlLy3  +hhakLy3  +hyjcLy3  +sdsjLy3  ;

        double sdsjMinustytm=tytmIn-tytmOut;
        double sdsjMinusty  =tyIn  -tyOut  ;
        double sdsjMinuslytm=lytmIn-lytmOut;
        double sdsjMinusly  =lyIn  -lyOut  ;

        deptUtils.outputMBI(Constant.BYJ,workbook,sheet,33,
                tytmIn,
                tyIn  ,
                lytmIn,
                lyIn  ,
                tytmOut,
                tyOut  ,
                lytmOut,
                lyOut  ,
                sdsjMinustytm,
                sdsjMinusty  ,
                sdsjMinuslytm,
                sdsjMinusly  );
    }

    /**
     * 招投标保证金
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:55
     */
    private void getOtherCash1(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入
        //监理
        //其他应收款-招投标保证金（贷方）
        double sljlTytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlTy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        //其他应付款-招投标保证金（贷方）
        double sljlTytm2=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlTy2  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLytm2=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLy2  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");

        //华海
        double hhakTytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");

        double hhakTytm2= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakTy2  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLytm2= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLy2  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");

        //恒远
        double hyjcTytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLytm1= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLy1  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");

        double hyjcTytm2= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcTy2  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLytm2= hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLy2  = hhakSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");

        //设计
        double sdsjTytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjTy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLytm1=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLy1  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");

        double sdsjTytm2=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjTy2  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLytm2=sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLy2  =sljlSingleCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");

        //合计
        double tytmIn = sljlTytm1+sljlTytm2+hhakTytm1+hhakTytm2+sdsjTytm1+sdsjTytm2+hyjcTytm1+hyjcTytm2;
        double tyIn   = sljlTy1  +sljlTy2  +hhakTy1  +hhakTy2  +sdsjTy1  +sdsjTy2  +hyjcTy1  +hyjcTy2  ;
        double lytmIn = sljlLytm1+sljlLytm2+hhakLytm1+hhakLytm2+sdsjLytm1+sdsjLytm2+hyjcLytm1+hyjcLytm2;
        double lyIn   = sljlLy1  +sljlLy2  +hhakLy1  +hhakLy2  +sdsjLy1  +sdsjLy2  +hyjcLy1  +hyjcLy2  ;

        //现金流出
        //监理
        //其他应收款-招投标保证金（借方）
        double sljlTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        //其他应付款-招投标保证金（借方）
        double sljlTytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlTy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");

        //华海
        double hhakTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");

        double hhakTytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakTy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");

        //恒远
        double hyjcTytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLytm3= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLy3  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");

        double hyjcTytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcTy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLytm4= hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLy4  = hhakSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");

        //设计
        double sdsjTytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjTy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLytm3=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLy3  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");

        double sdsjTytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjTy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLytm4=sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLy4  =sljlSingleDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");

        //合计
        double tytmOut=sljlTytm3+sljlTytm4+hhakTytm3+hhakTytm4+hyjcTytm3+hyjcTytm4+sdsjTytm3+sdsjTytm4;
        double tyOut  =sljlTy3  +sljlTy4  +hhakTy3  +hhakTy4  +hyjcTy3  +hyjcTy4  +sdsjTy3  +sdsjTy4  ;
        double lytmOut=sljlLytm3+sljlLytm4+hhakLytm3+hhakLytm4+hyjcLytm3+hyjcLytm4+sdsjLytm3+sdsjLytm4;
        double lyOut  =sljlLy3  +sljlLy4  +hhakLy3  +hhakLy4  +hyjcLy3  +hyjcLy4  +sdsjLy3  +sdsjLy4  ;

        //现金流量净额
        double sdsjMinustytm=tytmIn-tytmOut;
        double sdsjMinusty  =tyIn  -tyOut  ;
        double sdsjMinuslytm=lytmIn-lytmOut;
        double sdsjMinusly  =lyIn  -lyOut  ;

        deptUtils.outputMBI(Constant.ZTB,workbook,sheet,32,
                tytmIn,
                tyIn  ,
                lytmIn,
                lyIn  ,
                tytmOut,
                tyOut  ,
                lytmOut,
                lyOut  ,
                sdsjMinustytm,
                sdsjMinusty  ,
                sdsjMinuslytm,
                sdsjMinusly  );
    }

    /**
     * hhak单笔借记
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hhakSingleDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.singleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    /**
     * hhak单笔信贷
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hhakSingleCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                      String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.singleCreditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    /**
     * sljl单笔借方
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double sljlSingleDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                     String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.singleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount))+
                deptUtils.singleDeptSum(connection,commonSql.singleDebitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    /**
     * sljl单笔信贷
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double sljlSingleCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                      String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.singleCreditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount))+
                deptUtils.singleDeptSum(connection,commonSql.singleCreditAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    /**
     * 评估咨询
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getPgzxCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException{
        //现金流入：应收账款-开票（贷方）
        double[] pgzxInTytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInTy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInLytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInLy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");

        //主营业务成本
        double[] pgzxTytm= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxTy  = pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxLytm= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxLy  = pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");

        //应收账款-开票（借方）
        double[] pgzxInvoiceTytm = pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInvoiceTy   = pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInvoiceLytm = pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");
        double[] pgzxInvoiceLy   = pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"1122","112201");

        //现金流出
        double[] pgzxOutTytm=new double[pgzxInvoiceTytm.length];
        double[] pgzxOutTy  =new double[pgzxInvoiceTytm.length];
        double[] pgzxOutLytm=new double[pgzxInvoiceTytm.length];
        double[] pgzxOutLy  =new double[pgzxInvoiceTytm.length];
        for (int i = 0; i < pgzxInvoiceTytm.length; i++) {
            pgzxOutTytm[i]=pgzxTytm[i]+pgzxInvoiceTytm[i]/1.06*0.06+pgzxInvoiceTytm[i]/1.06*0.06*0.07+pgzxInvoiceTytm[i]/1.06*0.06*0.03+pgzxInvoiceTytm[i]/1.06*0.06*0.02;
            pgzxOutTy  [i]=pgzxTy  [i]+pgzxInvoiceTy  [i]/1.06*0.06+pgzxInvoiceTy  [i]/1.06*0.06*0.07+pgzxInvoiceTy  [i]/1.06*0.06*0.03+pgzxInvoiceTy  [i]/1.06*0.06*0.02;
            pgzxOutLytm[i]=pgzxLytm[i]+pgzxInvoiceLytm[i]/1.06*0.06+pgzxInvoiceLytm[i]/1.06*0.06*0.07+pgzxInvoiceLytm[i]/1.06*0.06*0.03+pgzxInvoiceLytm[i]/1.06*0.06*0.02;
            pgzxOutLy  [i]=pgzxLy  [i]+pgzxInvoiceLy  [i]/1.06*0.06+pgzxInvoiceLy  [i]/1.06*0.06*0.07+pgzxInvoiceLy  [i]/1.06*0.06*0.03+pgzxInvoiceLy  [i]/1.06*0.06*0.02;
        }

        //现金流量净额
        double[] pgzxMinusTytm =new double[pgzxInvoiceTytm.length];
        double[] pgzxMinusTy   =new double[pgzxInvoiceTytm.length];
        double[] pgzxMinusLytm =new double[pgzxInvoiceTytm.length];
        double[] pgzxMinusLy   =new double[pgzxInvoiceTytm.length];

        for (int i = 0; i < pgzxInvoiceTytm.length; i++) {
            pgzxMinusTytm[i]=pgzxInTytm[i]-pgzxOutTytm[i];
            pgzxMinusTy[i]  =pgzxInTy  [i]-pgzxOutTy  [i];
            pgzxMinusLytm[i]=pgzxInLytm[i]-pgzxOutLytm[i];
            pgzxMinusLy[i]  =pgzxInLy  [i]-pgzxOutLy  [i];
        }

        deptUtils.outputCFSExcel(Constant.PGZX,workbook,sheet,29,
                pgzxInTytm,
                pgzxInTy,
                pgzxInLytm,
                pgzxInLy  ,
                pgzxOutTytm,
                pgzxOutTy,
                pgzxOutLytm,
                pgzxOutLy,
                pgzxMinusTytm,
                pgzxMinusTy,
                pgzxMinusLytm,
                pgzxMinusLy);
    }

    /**
     * pgzx mbccredit
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double[] pgzxMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(Constant.SDSJ_DEPT_INCOME.get(4),connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    /**
     * pgzx mbcdebit1
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double[] pgzxMBCDebit1(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(Constant.SDSJ_DEPT_INCOME.get(4),connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    /**
     * pgzx mbcdebit
     *
     * @param connection     联系
     * @param commonSql      通用sql
     * @param deptUtils        dept sum
     * @param startDate      开始日期
     * @param endDate        结束日期
     * @param accountBook    账簿
     * @param accountBook1   账簿1
     * @param ledgerAccount  分类帐
     * @param ledgerAccount1 分类科目1
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double[] pgzxMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                  String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount,String ledgerAccount1) throws SQLException {
        return deptUtils.pgzxDeptSum1(connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1, ledgerAccount1));
    }

    /**
     * 石大设计
     */
    private void getSdsjCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double sdsjInTytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInTy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInLytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInLy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");

        //主营业务成本
        double sdsjTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");

        //应收账款-开票（借方）
        double sdsjInvoiceTytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoiceTy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoiceLytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoiceLy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");

        //现金流出
        double sdsjOutTytm=sdsjTytm+sdsjInvoiceTytm/1.06*0.06+sdsjInvoiceTytm/1.06*0.06*0.07+sdsjInvoiceTytm/1.06*0.06*0.03+sdsjInvoiceTytm/1.06*0.06*0.02;
        double sdsjOutTy  =sdsjTy  +sdsjInvoiceTy  /1.06*0.06+sdsjInvoiceTy  /1.06*0.06*0.07+sdsjInvoiceTy  /1.06*0.06*0.03+sdsjInvoiceTy  /1.06*0.06*0.02;
        double sdsjOutLytm=sdsjLytm+sdsjInvoiceLytm/1.06*0.06+sdsjInvoiceLytm/1.06*0.06*0.07+sdsjInvoiceLytm/1.06*0.06*0.03+sdsjInvoiceLytm/1.06*0.06*0.02;
        double sdsjOutLy  =sdsjLy  +sdsjInvoiceLy  /1.06*0.06+sdsjInvoiceLy  /1.06*0.06*0.07+sdsjInvoiceLy  /1.06*0.06*0.03+sdsjInvoiceLy  /1.06*0.06*0.02;

       //现金流量净额
        double sdsjMinusTytm=sdsjInTytm-sdsjOutTytm;
        double sdsjMinusTy  =sdsjInTy  -sdsjOutTy  ;
        double sdsjMinusLytm=sdsjInLytm-sdsjOutLytm;
        double sdsjMinusLy  =sdsjInLy  -sdsjOutLy  ;

        deptUtils.outputMBI(Constant.SDSJ,workbook,sheet,28,
                sdsjInTytm,
                sdsjInTy,
                sdsjInLytm,
                sdsjInLy,
                sdsjOutTytm,
                sdsjOutTy,
                sdsjOutLytm,
                sdsjOutLy,
                sdsjMinusTytm,
                sdsjMinusTy  ,
                sdsjMinusLytm,
                sdsjMinusLy  );

    }


    /**
     * 恒远检测
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getHyjcCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hyjcInTytm = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInTy   = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInLytm = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInLy   = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"1122","112201");

        //主营业务成本
        double hyjcTytm= hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcTy  = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcLytm= hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcLy  = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401");

        //应收账款-开票（借方）
        double hyjcInvoiceTytm = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInvoiceTy   = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInvoiceLytm = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"1122","112201");
        double hyjcInvoiceLy   = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"1122","112201");

        //现金流出
        double hyjcOutTytm=hyjcTytm+hyjcInvoiceTytm/1.06*0.06+hyjcInvoiceTytm/1.06*0.06*0.07+hyjcInvoiceTytm/1.06*0.06*0.03+hyjcInvoiceTytm/1.06*0.06*0.02;
        double hyjcOutTy  =hyjcTy  +hyjcInvoiceTy  /1.06*0.06+hyjcInvoiceTy  /1.06*0.06*0.07+hyjcInvoiceTy  /1.06*0.06*0.03+hyjcInvoiceTy  /1.06*0.06*0.02;
        double hyjcOutLytm=hyjcLytm+hyjcInvoiceLytm/1.06*0.06+hyjcInvoiceLytm/1.06*0.06*0.07+hyjcInvoiceLytm/1.06*0.06*0.03+hyjcInvoiceLytm/1.06*0.06*0.02;
        double hyjcOutLy  =hyjcLy  +hyjcInvoiceLy  /1.06*0.06+hyjcInvoiceLy  /1.06*0.06*0.07+hyjcInvoiceLy  /1.06*0.06*0.03+hyjcInvoiceLy  /1.06*0.06*0.02;

        double hyjcMinusTytm=hyjcInTytm-hyjcOutTytm;
        double hyjcMinusTy  =hyjcInTy  -hyjcOutTy  ;
        double hyjcMinusLytm=hyjcInLytm-hyjcOutLytm;
        double hyjcMinusLy  =hyjcInLy  -hyjcOutLy  ;

        deptUtils.outputMBI(Constant.HYJC,workbook,sheet,27,
                hyjcInTytm,
                hyjcInTy,
                hyjcInLytm,
                hyjcInLy,
                hyjcOutTytm,
                hyjcOutTy,
                hyjcOutLytm,
                hyjcOutLy,
                hyjcMinusTytm,
                hyjcMinusTy  ,
                hyjcMinusLytm,
                hyjcMinusLy  );

    }

    /**
     * hyjc信贷
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hyjcCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                              String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * hyjc借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hyjcDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                              String startDate,String endDate,String accountBook,String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount));
    }

    /**
     * hyjc借方
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hyjcDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * 华海安科
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getHhakCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hhakInTytm = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInTy   = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInLytm = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInLy   = hyjcCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"1122","112201");

        //主营业务成本
        double hhakTytm= hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double hhakTy  = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double hhakLytm= hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");
        double hhakLy  = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");

        //应收账款-开票（借方）
        double hhakInvoiceTytm = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInvoiceTy   = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInvoiceLytm = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double hhakInvoiceLy   = hyjcDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"1122","112201");

        //现金流出
        double hhakOutTytm=hhakTytm+hhakInvoiceTytm/1.06*0.06*1.12;
        double hhakOutTy  =hhakTy  +hhakInvoiceTy  /1.06*0.06*1.12;
        double hhakOutLytm=hhakLytm+hhakInvoiceLytm/1.06*0.06*1.12;
        double hhakOutLy  =hhakLy  +hhakInvoiceLy  /1.06*0.06*1.12;

        double hhakMinusTytm=hhakInTytm-hhakOutTytm;
        double hhakMinusTy  =hhakInTy  -hhakOutTy  ;
        double hhakMinusLytm=hhakInLytm-hhakOutLytm;
        double hhakMinusLy  =hhakInLy  -hhakOutLy  ;

        deptUtils.outputMBI(Constant.HHAK_DEPT_NAME.get(0),workbook,sheet,25,
                hhakInTytm,
                hhakInTy,
                hhakInLytm,
                hhakInLy,
                hhakOutTytm,
                hhakOutTy,
                hhakOutLytm,
                hhakOutLy,
                hhakMinusTytm,
                hhakMinusTy  ,
                hhakMinusLytm,
                hhakMinusLy  );
    }

    /**
     * 造价咨询
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getZjzxMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zjzxInTytm = zjzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInTy   = zjzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInLytm = zjzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInLy   = zjzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");

        //主营业务成本
        double zjzxTytm= zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxTy  = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxLytm= zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxLy  = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");

        //应收账款-开票（借方）
        double zjzxInvoiceTytm = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInvoiceTy   = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInvoiceLytm = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zjzxInvoiceLy   = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");

        //现金流出
        double zjzxOutTytm=zjzxTytm+zjzxInvoiceTytm/1.06*0.06+zjzxInvoiceTytm/1.06*0.06*0.07+zjzxInvoiceTytm/1.06*0.06*0.03+zjzxInvoiceTytm/1.06*0.06*0.02;
        double zjzxOutTy  =zjzxTy  +zjzxInvoiceTy  /1.06*0.06+zjzxInvoiceTy  /1.06*0.06*0.07+zjzxInvoiceTy  /1.06*0.06*0.03+zjzxInvoiceTy  /1.06*0.06*0.02;
        double zjzxOutLytm=zjzxLytm+zjzxInvoiceLytm/1.06*0.06+zjzxInvoiceLytm/1.06*0.06*0.07+zjzxInvoiceLytm/1.06*0.06*0.03+zjzxInvoiceLytm/1.06*0.06*0.02;
        double zjzxOutLy  =zjzxLy  +zjzxInvoiceLy  /1.06*0.06+zjzxInvoiceLy  /1.06*0.06*0.07+zjzxInvoiceLy  /1.06*0.06*0.03+zjzxInvoiceLy  /1.06*0.06*0.02;

        double zjzxMinusTytm=zjzxInTytm-zjzxOutTytm;
        double zjzxMinusTy  =zjzxInTy  -zjzxOutTy  ;
        double zjzxMinusLytm=zjzxInLytm-zjzxOutLytm;
        double zjzxMinusLy  =zjzxInLy  -zjzxOutLy  ;

        deptUtils.outputMBI(Constant.ZJZX_NAME,workbook,sheet,24,
                zjzxInTytm,
                zjzxInTy,
                zjzxInLytm,
                zjzxInLy,
                zjzxOutTytm,
                zjzxOutTy,
                zjzxOutLytm,
                zjzxOutLy,
                zjzxMinusTytm,
                zjzxMinusTy  ,
                zjzxMinusLytm,
                zjzxMinusLy  );
    }

    /**
     * zjzx信用
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zjzxCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                              String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * zjzx借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zjzxDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    /**
     * zjzx借方
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zjzxDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * 获取hhch-mbccash
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getHhchMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zbdlInTytm = hhchCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");
        double zbdlInTy   = hhchCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"1122","112201");

        //主营业务成本
        double zbdlTytm= hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double zbdlTy  = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double zbdlLytm= hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");
        double zbdlLy  = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");

        //主营业务收入-开票（借方）
        double zbdlInvoiceTytm = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"6001","600101");
        double zbdlInvoiceTy   = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),"6001","600101");
        double zbdlInvoiceLytm = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"6001","600101");
        double zbdlInvoiceLy   = hhchDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),"6001","600101");

        //现金流出
        double zbdlOutTytm=zbdlTytm+zbdlInvoiceTytm*0.06+zbdlInvoiceTytm*0.06*0.07+zbdlInvoiceTytm*0.06*0.03+zbdlInvoiceTytm*0.06*0.02;
        double zbdlOutTy  =zbdlTy  +zbdlInvoiceTy  *0.06+zbdlInvoiceTy  *0.06*0.07+zbdlInvoiceTy  *0.06*0.03+zbdlInvoiceTy  *0.06*0.02;
        double zbdlOutLytm=zbdlLytm+zbdlInvoiceLytm*0.06+zbdlInvoiceLytm*0.06*0.07+zbdlInvoiceLytm*0.06*0.03+zbdlInvoiceLytm*0.06*0.02;
        double zbdlOutLy  =zbdlLy  +zbdlInvoiceLy  *0.06+zbdlInvoiceLy  *0.06*0.07+zbdlInvoiceLy  *0.06*0.03+zbdlInvoiceLy  *0.06*0.02;

        double zbdlMinusTytm=zbdlInTytm-zbdlOutTytm;
        double zbdlMinusTy  =zbdlInTy  -zbdlOutTy  ;


        deptUtils.outputMBI(Constant.HHAK_DEPT_NAME.get(5),workbook,sheet,26,
                zbdlInTytm,
                zbdlInTy,
                0.0,
                0.0,
                zbdlOutTytm,
                zbdlOutTy,
                zbdlOutLytm,
                zbdlOutLy,
                zbdlMinusTytm,
                zbdlMinusTy  ,
                0.0,
                0.0  );
    }

    /**
     * hhch信用
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hhchCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                              String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.HHAK_DEPT.get(3),commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * hhch借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hhchDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.HHAK_DEPT.get(3),commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    /**
     * hhch借方
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double hhchDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.HHAK_DEPT.get(3),commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * 招标代理
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getZbdlMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zbdlInTytm = zbdlCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInTy   = zbdlCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInLytm = zbdlCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInLy   = zbdlCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");

        //主营业务成本
        double zbdlTytm= zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlTy  = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlLytm= zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlLy  = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");

        //应收账款-开票（借方）
        double zbdlInvoiceTytm = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInvoiceTy   = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInvoiceLytm = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");
        double zbdlInvoiceLy   = zbdlDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"1122","112201");

        //现金流出
        double zbdlOutTytm=zbdlTytm+zbdlInvoiceTytm/1.06*0.06+zbdlInvoiceTytm/1.06*0.06*0.07+zbdlInvoiceTytm/1.06*0.06*0.03+zbdlInvoiceTytm/1.06*0.06*0.02;
        double zbdlOutTy  =zbdlTy  +zbdlInvoiceTy  /1.06*0.06+zbdlInvoiceTy  /1.06*0.06*0.07+zbdlInvoiceTy  /1.06*0.06*0.03+zbdlInvoiceTy  /1.06*0.06*0.02;
        double zbdlOutLytm=zbdlLytm+zbdlInvoiceLytm/1.06*0.06+zbdlInvoiceLytm/1.06*0.06*0.07+zbdlInvoiceLytm/1.06*0.06*0.03+zbdlInvoiceLytm/1.06*0.06*0.02;
        double zbdlOutLy  =zbdlLy  +zbdlInvoiceLy  /1.06*0.06+zbdlInvoiceLy  /1.06*0.06*0.07+zbdlInvoiceLy  /1.06*0.06*0.03+zbdlInvoiceLy  /1.06*0.06*0.02;

        double zbdlMinusTytm=zbdlInTytm-zbdlOutTytm;
        double zbdlMinusTy  =zbdlInTy  -zbdlOutTy  ;
        double zbdlMinusLytm=zbdlInLytm-zbdlOutLytm;
        double zbdlMinusLy  =zbdlInLy  -zbdlOutLy  ;

        deptUtils.outputMBI(Constant.ZBDL_NAME,workbook,sheet,23,
                zbdlInTytm,
                zbdlInTy,
                zbdlInLytm,
                zbdlInLy,
                zbdlOutTytm,
                zbdlOutTy,
                zbdlOutLytm,
                zbdlOutLy,
                zbdlMinusTytm,
                zbdlMinusTy  ,
                zbdlMinusLytm,
                zbdlMinusLy  );
    }

    /**
     * zbdl信贷
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zbdlCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * zbdl借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zbdlDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    /**
     * zbdl借方
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double zbdlDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                              String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    /**
     * 监理-生产部门
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private void getSljlMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double[] sljlMBCInTytm = sljlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInTy   = sljlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInLytm = sljlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInLy   = sljlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");

        //主营业务成本
        double[] sljlMBCTytm= sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCTy  = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCLytm= sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCLy  = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");

        //应收账款-开票（借方）
        double[] sljlMBCInvoiceTytm = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInvoiceTy   = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInvoiceLytm = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");
        double[] sljlMBCInvoiceLy   = sljlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"1122","112201");

        //现金流出
        double[] sljlMBCOutTytm=new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCOutTy  =new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCOutLytm=new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCOutLy  =new double[sljlMBCInvoiceTytm.length];
        for (int i = 0; i < sljlMBCInvoiceTytm.length; i++) {
            sljlMBCOutTytm[i]=sljlMBCTytm[i]+sljlMBCInvoiceTytm[i]/1.06*0.06+sljlMBCInvoiceTytm[i]/1.06*0.06*0.07+sljlMBCInvoiceTytm[i]/1.06*0.06*0.03+sljlMBCInvoiceTytm[i]/1.06*0.06*0.02;
            sljlMBCOutTy  [i]=sljlMBCTy  [i]+sljlMBCInvoiceTy  [i]/1.06*0.06+sljlMBCInvoiceTy  [i]/1.06*0.06*0.07+sljlMBCInvoiceTy  [i]/1.06*0.06*0.03+sljlMBCInvoiceTy  [i]/1.06*0.06*0.02;
            sljlMBCOutLytm[i]=sljlMBCLytm[i]+sljlMBCInvoiceLytm[i]/1.06*0.06+sljlMBCInvoiceLytm[i]/1.06*0.06*0.07+sljlMBCInvoiceLytm[i]/1.06*0.06*0.03+sljlMBCInvoiceLytm[i]/1.06*0.06*0.02;
            sljlMBCOutLy  [i]=sljlMBCLy  [i]+sljlMBCInvoiceLy  [i]/1.06*0.06+sljlMBCInvoiceLy  [i]/1.06*0.06*0.07+sljlMBCInvoiceLy  [i]/1.06*0.06*0.03+sljlMBCInvoiceLy  [i]/1.06*0.06*0.02;
        }

        //现金流量净额
        double[] sljlMBCMinusTytm =new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCMinusTy   =new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCMinusLytm =new double[sljlMBCInvoiceTytm.length];
        double[] sljlMBCMinusLy   =new double[sljlMBCInvoiceTytm.length];

        for (int i = 0; i < sljlMBCInvoiceTytm.length; i++) {
            sljlMBCMinusTytm[i]=sljlMBCInTytm[i]-sljlMBCOutTytm[i];
            sljlMBCMinusTy[i]  =sljlMBCInTy  [i]-sljlMBCOutTy  [i];
            sljlMBCMinusLytm[i]=sljlMBCInLytm[i]-sljlMBCOutLytm[i];
            sljlMBCMinusLy[i]  =sljlMBCInLy  [i]-sljlMBCOutLy  [i];
        }

        deptUtils.outputCFSExcel(Constant.SLJL_BRANCH_OFFICE_NAME,workbook,sheet,13,
                sljlMBCInTytm,
                sljlMBCInTy,
                sljlMBCInLytm,
                sljlMBCInLy  ,
                sljlMBCOutTytm,
                sljlMBCOutTy,
                sljlMBCOutLytm,
                sljlMBCOutLy,
                sljlMBCMinusTytm,
                sljlMBCMinusTy,
                sljlMBCMinusLytm,
                sljlMBCMinusLy);

    }

    /**
     * sljl mbcdebit
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param accountBook1  账簿1
     * @param ledgerAccount 分类帐
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:56
     */
    private double[] sljlMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate,String endDate,String accountBook,String accountBook1,String ledgerAccount) throws SQLException {
        return deptUtils.sljlManufactureDeptSum(connection,
                commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount),
                commonSql.debitAmount(startDate,endDate,accountBook1,ledgerAccount));
    }

    /**
     * sljl mbcdebit
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:57
     */
    private double[] sljlMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                  String startDate,String endDate,String accountBook,String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.sljlManufactureDeptSum(connection,
                commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    /**
     * sljl mbccredit
     *
     * @param connection      联系
     * @param commonSql       通用sql
     * @param deptUtils         dept sum
     * @param startDate       开始日期
     * @param endDate         结束日期
     * @param accountBook     账簿
     * @param accountBook1    账簿1
     * @param ledgerAccount   分类帐
     * @param specificAccount 特定帐户
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:57
     */
    private double[] sljlMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate,String endDate,String accountBook,String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.sljlManufactureDeptSum(connection,
                commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    /**
     * 监理-管理部门
     *
     * @param connection 联系
     * @param workbook   工作簿
     * @param sheet      床单
     * @param dateUtils  日期实用程序
     * @param commonSql  通用sql
     * @param deptUtils    dept sum
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:57
     */
    private void getSljlOCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //监理管理费用
        double[] sljlOverheadTytm= sljlOverheadDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadTy  = sljlOverheadDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadLytm= sljlOverheadDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadLy  = sljlOverheadDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), "6602");

        //监理销售费用
        double sljlSellingTytm= singleDeptDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingTy  = singleDeptDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingLytm= singleDeptDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingLy  = singleDeptDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601");

        sljlOverheadTytm[0]+=sljlSellingTytm;
        sljlOverheadTy  [0]+=sljlSellingTy  ;
        sljlOverheadLytm[0]+=sljlSellingLytm;
        sljlOverheadLy  [0]+=sljlSellingLy  ;

        deptUtils.outputMBExel(Constant.SLJL_MANAGE_DEPT_NAME,workbook,sheet,6,8,
                sljlOverheadTytm,
                sljlOverheadTy,
                sljlOverheadLytm,
                sljlOverheadLy  ,
                sljlSellingTytm,
                sljlSellingTy,
                sljlSellingLytm,
                sljlSellingLy  );
    }

    /**
     * 链路间接费用借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:57
     */
    private double[] sljlOverheadDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.sljlOverheadDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    /**
     * 单部门借方
     *
     * @param connection    联系
     * @param commonSql     通用sql
     * @param deptUtils       dept sum
     * @param startDate     开始日期
     * @param endDate       结束日期
     * @param accountBook   账簿
     * @param ledgerAccount 分类帐
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 13:59:57
     */
    private double singleDeptDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }
}
