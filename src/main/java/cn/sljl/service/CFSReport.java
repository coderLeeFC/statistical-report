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

    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        DateUtils dateUtils = new DateUtils();
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();

        titleModel.createCFSTitle(workbook, sheet, 20,Constant.CFS_TITLE_RIGHT,Constant.TITLE_SIXTH_RIGHT_CFS);

        getSljlOCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getSljlMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZbdlCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZjzxCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHhakCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHhchMBCCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHyjcCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getSdsjCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getPgzxCash(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash1(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash2(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getOtherCash3(connection,workbook,sheet,dateUtils,commonSql, deptUtils);

    }

    private void getOtherCash3(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //监理
        //企业所得税
        double sljlTytm1=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlTy1  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlLytm1=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        double sljlLy1  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222104");
        //个人所得税
        double sljlTytm2=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlTy2  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlLytm2=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        double sljlLy2  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222105");
        //房产税
        double sljlTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        double sljlLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222106");
        //城镇土地使用税
        double sljlTytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlTy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlLytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        double sljlLy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222107");
        //印花税
        double sljlTytm5=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlTy5  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlLytm5=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");
        double sljlLy5  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2221","222108");

        //华海
        double hhakTytm1= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakTy1  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakLytm1= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakLy1  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222104");
        double hhakTytm2= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakTy2  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakLytm2= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakLy2  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222105");
        double hhakTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222106");
        double hhakTytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakTy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakLytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakLy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222107");
        double hhakTytm5= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakTy5  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakLytm5= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222108");
        double hhakLy5  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2221","222108");

        //恒远
        double hyjcTytm1= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcTy1  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcLytm1= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcLy1  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222104");
        double hyjcTytm2= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcTy2  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcLytm2= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcLy2  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222105");
        double hyjcTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222106");
        double hyjcTytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcTy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcLytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcLy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222107");
        double hyjcTytm5= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcTy5  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcLytm5= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222108");
        double hyjcLy5  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2221","222108");

        //设计
        double sdsjTytm1=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjTy1  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjLytm1=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjLy1  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222104");
        double sdsjTytm2=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjTy2  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjLytm2=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjLy2  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222105");
        double sdsjTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222106");
        double sdsjTytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjTy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjLytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjLy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222107");
        double sdsjTytm5=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjTy5  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjLytm5=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");
        double sdsjLy5  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2221","222108");

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

    private void getOtherCash2(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入【其他应收款-个人（贷方）】
        //监理
        double sljlTytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlTy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        //华海
        double hhakTytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakTy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        //恒远
        double hyjcTytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcTy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        //设计
        double sdsjTytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjTy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");

        double tytmIn = sljlTytm1+hhakTytm1+sdsjTytm1+hyjcTytm1;
        double tyIn   = sljlTy1  +hhakTy1  +sdsjTy1  +hyjcTy1  ;
        double lytmIn = sljlLytm1+hhakLytm1+sdsjLytm1+hyjcLytm1;
        double lyIn   = sljlLy1  +hhakLy1  +sdsjLy1  +hyjcLy1  ;

        //现金流出【其他应收款-个人（借方）】
        //监理
        double sljlTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        double sljlLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123102");
        //华海
        double hhakTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        double hhakLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123102");
        //恒远
        double hyjcTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        double hyjcLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123102");
        //设计
        double sdsjTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");
        double sdsjLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123102");

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

    private void getOtherCash1(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入
        //监理
        //其他应收款-招投标保证金（贷方）
        double sljlTytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlTy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        //其他应付款-招投标保证金（贷方）
        double sljlTytm2=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlTy2  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLytm2=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLy2  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");

        //华海
        double hhakTytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTytm2= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakTy2  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLytm2= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLy2  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");

        //恒远
        double hyjcTytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLytm1= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLy1  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTytm2= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcTy2  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLytm2= deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLy2  = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");

        //设计
        double sdsjTytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjTy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLytm1=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLy1  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");

        double sdsjTytm2=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjTy2  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLytm2=deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLy2  =deptUtils.singleDeptCredit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");

        //合计
        double tytmIn = sljlTytm1+sljlTytm2+hhakTytm1+hhakTytm2+sdsjTytm1+sdsjTytm2+hyjcTytm1+hyjcTytm2;
        double tyIn   = sljlTy1  +sljlTy2  +hhakTy1  +hhakTy2  +sdsjTy1  +sdsjTy2  +hyjcTy1  +hyjcTy2  ;
        double lytmIn = sljlLytm1+sljlLytm2+hhakLytm1+hhakLytm2+sdsjLytm1+sdsjLytm2+hyjcLytm1+hyjcLytm2;
        double lyIn   = sljlLy1  +sljlLy2  +hhakLy1  +hhakLy2  +sdsjLy1  +sdsjLy2  +hyjcLy1  +hyjcLy2  ;

        //现金流出
        //监理
        //其他应收款-招投标保证金（借方）
        double sljlTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        double sljlLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "1231","123103");
        //其他应付款-招投标保证金（借方）
        double sljlTytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlTy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");
        double sljlLy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "2241","224103");

        //华海
        double hhakTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "1231","123103");
        double hhakTytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakTy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");
        double hhakLy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "2241","224103");

        //恒远
        double hyjcTytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLytm3= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcLy3  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "1231","123103");
        double hyjcTytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcTy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLytm4= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");
        double hyjcLy4  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "2241","224103");

        //设计
        double sdsjTytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjTy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLytm3=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");
        double sdsjLy3  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1231","123103");

        double sdsjTytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjTy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLytm4=deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");
        double sdsjLy4  =deptUtils.singleDeptDebit1(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "2241","224103");

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

    private void getPgzxCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException{
        //现金流入：应收账款-开票（贷方）
        double[] pgzxInTytm = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInTy   = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInLytm = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInLy   = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double[] pgzxTytm= deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxTy  = deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxLytm= deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxLy  = deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);

        //应收账款-开票（借方）
        double[] pgzxInvoiceTytm = deptUtils.pgzxDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoiceTy   = deptUtils.pgzxDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoiceLytm = deptUtils.pgzxDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoiceLy   = deptUtils.pgzxDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getSdsjCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double sdsjInTytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInTy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInLytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInLy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double sdsjTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);

        //应收账款-开票（借方）
        double sdsjInvoiceTytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoiceTy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoiceLytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoiceLy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getHyjcCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hyjcInTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInTy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInLy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double hyjcTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);

        //应收账款-开票（借方）
        double hyjcInvoiceTytm = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInvoiceTy   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInvoiceLytm = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInvoiceLy   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getHhchMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zbdlInTytm = deptUtils.singleDeptCredit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInTy   = deptUtils.singleDeptCredit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double zbdlTytm= deptUtils.singleDeptDebit(Constant.HHAK_DEPT.get(3),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double zbdlTy  = deptUtils.singleDeptDebit(Constant.HHAK_DEPT.get(3),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double zbdlLytm= deptUtils.singleDeptDebit(Constant.HHAK_DEPT.get(3),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double zbdlLy  = deptUtils.singleDeptDebit(Constant.HHAK_DEPT.get(3),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);

        //主营业务收入-开票（借方）
        double zbdlInvoiceTytm = deptUtils.singleDeptDebit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoiceTy   = deptUtils.singleDeptDebit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoiceLytm = deptUtils.singleDeptDebit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoiceLy   = deptUtils.singleDeptDebit(connection, commonSql, Constant.HHAK_DEPT.get(3), dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

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

    private void getHhakCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hhakInTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInTy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInLy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double hhakTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double hhakTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double hhakLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double hhakLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);

        //管理费用
        double hhakOCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.OVERHEAD);
        double hhakOCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.OVERHEAD);
        double hhakOCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.OVERHEAD);
        double hhakOCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.OVERHEAD);

        //销售费用
        double hhakSCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.SELLING_EXPENSES);
        double hhakSCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.SELLING_EXPENSES);
        double hhakSCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.SELLING_EXPENSES);
        double hhakSCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.SELLING_EXPENSES);

        //应收账款-开票（借方）
        double hhakInvoiceTytm = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInvoiceTy   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInvoiceLytm = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInvoiceLy   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //现金流出
        double hhakOutTytm=hhakOCTytm+hhakSCTytm+hhakTytm+hhakInvoiceTytm/1.06*0.06*1.12;
        double hhakOutTy  =hhakOCTy  +hhakSCTy  +hhakTy  +hhakInvoiceTy  /1.06*0.06*1.12;
        double hhakOutLytm=hhakOCLytm+hhakSCLytm+hhakLytm+hhakInvoiceLytm/1.06*0.06*1.12;
        double hhakOutLy  =hhakOCLy  +hhakSCLy  +hhakLy  +hhakInvoiceLy  /1.06*0.06*1.12;

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

    private void getZjzxCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zjzxInTytm = deptUtils.singleDeptCredit(connection, commonSql, Constant.ZJZX, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInTy   = deptUtils.singleDeptCredit(connection, commonSql, Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInLytm = deptUtils.singleDeptCredit(connection, commonSql, Constant.ZJZX, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInLy   = deptUtils.singleDeptCredit(connection, commonSql, Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double zjzxTytm= deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxTy  = deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxLytm= deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxLy  = deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);

        //应收账款-开票（借方）
        double zjzxInvoiceTytm = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInvoiceTy   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInvoiceLytm = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInvoiceLy   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getZbdlCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zbdlInTytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInTy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInLytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInLy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double zbdlTytm= deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlTy  = deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlLytm= deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlLy  = deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);

        //应收账款-开票（借方）
        double zbdlInvoiceTytm = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInvoiceTy   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInvoiceLytm = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInvoiceLy   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getSljlMBCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double[] sljlMBCInTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInTy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInLy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double[] sljlMBCTytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCTy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCLytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCLy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);

        //应收账款-开票（借方）
        double[] sljlMBCInvoiceTytm = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInvoiceTy   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInvoiceLytm = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInvoiceLy   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

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

    private void getSljlOCCash(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //监理管理费用
        double[] sljlOverheadTytm= deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), Constant.OVERHEAD);
        double[] sljlOverheadTy  = deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), Constant.OVERHEAD);
        double[] sljlOverheadLytm= deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), Constant.OVERHEAD);
        double[] sljlOverheadLy  = deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), Constant.OVERHEAD);

        //监理销售费用
        double sljlSellingTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.SELLING_EXPENSES);
        double sljlSellingTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.SELLING_EXPENSES);
        double sljlSellingLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.SELLING_EXPENSES);
        double sljlSellingLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.SELLING_EXPENSES);

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

}
