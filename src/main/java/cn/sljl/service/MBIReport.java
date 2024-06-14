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
 * 营业收入&营业总成本&毛利率
 * @author wangeqiu
 * @version 1.0
 * @date 2024/3/16 13:42
 */
public class MBIReport {

    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        DateUtils dateUtils = new DateUtils();
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();

        titleModel.createMBTitle(workbook, sheet,24);

        getSljlOverhead(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getSljlMBC(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getZbdl(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getZjzx(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getHhak(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getHyjc(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getSdsj(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
        getPgzx(connection, workbook, sheet, dateUtils,commonSql, deptUtils);
    }

    private void getPgzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] pgzxRevenueRecognitionTytm = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME);
        double[] pgzxRevenueRecognitionTy   = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME);
        double[] pgzxRevenueRecognitionLytm = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME);
        double[] pgzxRevenueRecognitionLy   = deptUtils.pgzxCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double[] pgzxInvoicingRevenueTytm = deptUtils.pgzx(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING,Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoicingRevenueTy   = deptUtils.pgzx(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING,Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoicingRevenueLytm = deptUtils.pgzx(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING,Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInvoicingRevenueLy   = deptUtils.pgzx(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING,Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //主营业务成本
        double[] pgzxMBCTytm= deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxMBCTy  = deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxMBCLytm= deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);
        double[] pgzxMBCLy  = deptUtils.pgzx(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), Constant.MAIN_BUSINESS_COST,Constant.INVENTORY);

        //管理费用（设计）
        double pgzxOCTytm= deptUtils.pgzxOCDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), Constant.OVERHEAD);
        double pgzxOCTy  = deptUtils.pgzxOCDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), Constant.OVERHEAD);
        double pgzxOCLytm= deptUtils.pgzxOCDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), Constant.OVERHEAD);
        double pgzxOCLy  = deptUtils.pgzxOCDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), Constant.OVERHEAD);

        //设计总成本
        pgzxMBCTytm[2]+=pgzxOCTytm;
        pgzxMBCTy  [2]+=pgzxOCTy  ;
        pgzxMBCLytm[2]+=pgzxOCLytm;
        pgzxMBCLy  [2]+=pgzxOCLy  ;

        //营业总成本
        pgzxMBCTytm[0]+=pgzxOCTytm;
        pgzxMBCTy  [0]+=pgzxOCTy  ;
        pgzxMBCLytm[0]+=pgzxOCLytm;
        pgzxMBCLy  [0]+=pgzxOCLy  ;

        //毛利率
        double[] pgzxGrossMarginTytm=new double[pgzxMBCTytm.length];
        double[] pgzxGrossMarginTy  =new double[pgzxMBCTytm.length];
        double[] pgzxGrossMarginLytm=new double[pgzxMBCTytm.length];
        double[] pgzxGrossMarginLy  =new double[pgzxMBCTytm.length];

        for (int i = 0; i < pgzxMBCTytm.length; i++) {
            pgzxGrossMarginTytm[i]=(pgzxRevenueRecognitionTytm[i]-pgzxMBCTytm[i])/pgzxRevenueRecognitionTytm[i]*100;
            pgzxGrossMarginTy  [i]=(pgzxRevenueRecognitionTy  [i]-pgzxMBCTy  [i])/pgzxRevenueRecognitionTy  [i]*100;
            pgzxGrossMarginLytm[i]=(pgzxRevenueRecognitionLytm[i]-pgzxMBCLytm[i])/pgzxRevenueRecognitionLytm[i]*100;
            pgzxGrossMarginLy  [i]=(pgzxRevenueRecognitionLy  [i]-pgzxMBCLy  [i])/pgzxRevenueRecognitionLy  [i]*100;
        }

        deptUtils.outputMB1Excel(Constant.PGZX,workbook,sheet,33,
                pgzxRevenueRecognitionTytm,
                pgzxRevenueRecognitionTy,
                pgzxRevenueRecognitionLytm,
                pgzxRevenueRecognitionLy  ,
                pgzxInvoicingRevenueTytm,
                pgzxInvoicingRevenueTy,
                pgzxInvoicingRevenueLytm,
                pgzxInvoicingRevenueLy  ,
                pgzxMBCTytm,
                pgzxMBCTy,
                pgzxMBCLytm,
                pgzxMBCLy  ,
                pgzxGrossMarginTytm,
                pgzxGrossMarginTy,
                pgzxGrossMarginLytm,
                pgzxGrossMarginLy  );
    }

    private void getSdsj(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double sdsjRevenueRecognitionTytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.MAIN_BUSINESS_INCOME);
        double sdsjRevenueRecognitionTy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.MAIN_BUSINESS_INCOME);
        double sdsjRevenueRecognitionLytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.MAIN_BUSINESS_INCOME);
        double sdsjRevenueRecognitionLy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double sdsjInvoicingRevenueTytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoicingRevenueTy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoicingRevenueLytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInvoicingRevenueLy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //存货
        double sdsjMBCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjMBCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjMBCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);
        double sdsjMBCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.INVENTORY);

        //管理费用
        double sdsjOCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.OVERHEAD);
        double sdsjOCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.OVERHEAD);
        double sdsjOCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.OVERHEAD);
        double sdsjOCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.OVERHEAD);

        //销售费用
        double sdsjSCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.SELLING_EXPENSES);
        double sdsjSCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.SELLING_EXPENSES);
        double sdsjSCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.SELLING_EXPENSES);
        double sdsjSCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), Constant.SELLING_EXPENSES);

        double sdsjDevelopTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), Constant.RD_EXPENSES);
        double sdsjDevelopTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), Constant.RD_EXPENSES);
        double sdsjDevelopLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), Constant.RD_EXPENSES);
        double sdsjDevelopLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), Constant.RD_EXPENSES);


        //营业总成本
        double sdsjCostTytm=sdsjMBCTytm+sdsjOCTytm+sdsjSCTytm+sdsjDevelopTytm;
        double sdsjCostTy  =sdsjMBCTy  +sdsjOCTy  +sdsjSCTy  +sdsjDevelopTy  ;
        double sdsjCostLytm=sdsjMBCLytm+sdsjOCLytm+sdsjSCLytm+sdsjDevelopLytm;
        double sdsjCostLy  =sdsjMBCLy  +sdsjOCLy  +sdsjSCLy  +sdsjDevelopLy  ;

        //毛利率
        double sdsjGrossMarginTytm=(sdsjRevenueRecognitionTytm-sdsjCostTytm)/sdsjRevenueRecognitionTytm*100;
        double sdsjGrossMarginTy  =(sdsjRevenueRecognitionTy  -sdsjCostTy  )/sdsjRevenueRecognitionTy  *100;
        double sdsjGrossMarginLytm=(sdsjRevenueRecognitionLytm-sdsjCostLytm)/sdsjRevenueRecognitionLytm*100;
        double sdsjGrossMarginLy  =(sdsjRevenueRecognitionLy  -sdsjCostLy  )/sdsjRevenueRecognitionLy  *100;

        deptUtils.outputMB1Excel(Constant.SDSJ,workbook,sheet,32,
                sdsjRevenueRecognitionTytm,
                sdsjRevenueRecognitionTy,
                sdsjRevenueRecognitionLytm,
                sdsjRevenueRecognitionLy,
                sdsjInvoicingRevenueTytm,
                sdsjInvoicingRevenueTy,
                sdsjInvoicingRevenueLytm,
                sdsjInvoicingRevenueLy,
                sdsjCostTytm,
                sdsjCostTy  ,
                sdsjCostLytm,
                sdsjCostLy  ,
                sdsjGrossMarginTytm,
                sdsjGrossMarginTy,
                sdsjGrossMarginLytm,
                sdsjGrossMarginLy);
    }

    private void getHyjc(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double hyjcRevenueRecognitionTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME);
        double hyjcRevenueRecognitionTy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME);
        double hyjcRevenueRecognitionLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME);
        double hyjcRevenueRecognitionLy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double hyjcInvoicingRevenueTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double hyjcInvoicingRevenueTy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double hyjcInvoicingRevenueLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double hyjcInvoicingRevenueLy   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

        //主营业务成本
        double hyjcMBCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcMBCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcMBCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);
        double hyjcMBCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.MAIN_BUSINESS_COST);

        //管理费用
        double hyjcOCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.OVERHEAD);
        double hyjcOCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.OVERHEAD);
        double hyjcOCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.OVERHEAD);
        double hyjcOCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.OVERHEAD);

        //销售费用
        double hyjcSCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.SELLING_EXPENSES);
        double hyjcSCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), Constant.SELLING_EXPENSES);
        double hyjcSCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.SELLING_EXPENSES);
        double hyjcSCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), Constant.SELLING_EXPENSES);

        //营业总成本
        double hyjcCostTytm=hyjcMBCTytm+hyjcOCTytm+hyjcSCTytm;
        double hyjcCostTy  =hyjcMBCTy  +hyjcOCTy  +hyjcSCTy  ;
        double hyjcCostLytm=hyjcMBCLytm+hyjcOCLytm+hyjcSCLytm;
        double hyjcCostLy  =hyjcMBCLy  +hyjcOCLy  +hyjcSCLy  ;

        //毛利率
        double hyjcGrossMarginTytm=(hyjcRevenueRecognitionTytm-hyjcCostTytm)/hyjcRevenueRecognitionTytm*100;
        double hyjcGrossMarginTy  =(hyjcRevenueRecognitionTy  -hyjcCostTy  )/hyjcRevenueRecognitionTy  *100;
        double hyjcGrossMarginLytm=(hyjcRevenueRecognitionLytm-hyjcCostLytm)/hyjcRevenueRecognitionLytm*100;
        double hyjcGrossMarginLy  =(hyjcRevenueRecognitionLy  -hyjcCostLy  )/hyjcRevenueRecognitionLy  *100;

        deptUtils.outputMB1Excel(Constant.HYJC,workbook,sheet,31,
                hyjcRevenueRecognitionTytm,
                hyjcRevenueRecognitionTy,
                hyjcRevenueRecognitionLytm,
                hyjcRevenueRecognitionLy,
                hyjcInvoicingRevenueTytm,
                hyjcInvoicingRevenueTy,
                hyjcInvoicingRevenueLytm,
                hyjcInvoicingRevenueLy,
                hyjcCostTytm,
                hyjcCostTy  ,
                hyjcCostLytm,
                hyjcCostLy  ,
                hyjcGrossMarginTytm,
                hyjcGrossMarginTy,
                hyjcGrossMarginLytm,
                hyjcGrossMarginLy);
    }

    private void getHhak(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] hhakRevenueRecognitionTytm= deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME);
        double[] hhakRevenueRecognitionTy  = deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME);
        double[] hhakRevenueRecognitionLytm= deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME);
        double[] hhakRevenueRecognitionLy  = deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double[] hhakInvoicingRevenueTytm= deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] hhakInvoicingRevenueTy  = deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] hhakInvoicingRevenueLytm= deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] hhakInvoicingRevenueLy  = deptUtils.hhakCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

        //主营业务成本
        double[] hhakMBCTytm= deptUtils.hhakDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double[] hhakMBCTy  = deptUtils.hhakDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double[] hhakMBCLytm= deptUtils.hhakDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);
        double[] hhakMBCLy  = deptUtils.hhakDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), Constant.MAIN_BUSINESS_COST);

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

        //营业总成本
        hhakMBCTytm[0]+=(hhakOCTytm+hhakSCTytm);
        hhakMBCTy  [0]+=(hhakOCTy  +hhakSCTy  );
        hhakMBCLytm[0]+=(hhakOCLytm+hhakSCLytm);
        hhakMBCLy  [0]+=(hhakOCLy  +hhakSCLy  );

        //北京营业总成本
        hhakMBCTytm[1]+=(hhakOCTytm+hhakSCTytm);
        hhakMBCTy  [1]+=(hhakOCTy  +hhakSCTy  );
        hhakMBCLytm[1]+=(hhakOCLytm+hhakSCLytm);
        hhakMBCLy  [1]+=(hhakOCLy  +hhakSCLy  );

        //毛利率
        double[] hhakGrossMarginTytm=new double[hhakMBCTytm.length];
        double[] hhakGrossMarginTy  =new double[hhakMBCTytm.length];
        double[] hhakGrossMarginLytm=new double[hhakMBCTytm.length];
        double[] hhakGrossMarginLy  =new double[hhakMBCTytm.length];

        for (int i = 0; i < hhakMBCTytm.length; i++) {
            hhakGrossMarginTytm[i]=(hhakRevenueRecognitionTytm[i]-hhakMBCTytm[i])/hhakRevenueRecognitionTytm[i]*100;
            hhakGrossMarginTy  [i]=(hhakRevenueRecognitionTy  [i]-hhakMBCTy  [i])/hhakRevenueRecognitionTy  [i]*100;
            hhakGrossMarginLytm[i]=(hhakRevenueRecognitionLytm[i]-hhakMBCLytm[i])/hhakRevenueRecognitionLytm[i]*100;
            hhakGrossMarginLy  [i]=(hhakRevenueRecognitionLy  [i]-hhakMBCLy  [i])/hhakRevenueRecognitionLy  [i]*100;
        }

        deptUtils.outputMB1Excel(Constant.HHAK_DEPT_NAME,workbook,sheet,25,
                hhakRevenueRecognitionTytm,
                hhakRevenueRecognitionTy,
                hhakRevenueRecognitionLytm,
                hhakRevenueRecognitionLy  ,
                hhakInvoicingRevenueTytm,
                hhakInvoicingRevenueTy,
                hhakInvoicingRevenueLytm,
                hhakInvoicingRevenueLy  ,
                hhakMBCTytm,
                hhakMBCTy,
                hhakMBCLytm,
                hhakMBCLy  ,
                hhakGrossMarginTytm,
                hhakGrossMarginTy,
                hhakGrossMarginLytm,
                hhakGrossMarginLy  ,
                hhakOCTytm,
                hhakOCTy,
                hhakOCLytm,
                hhakOCLy  );
    }

    private void getZjzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double zjzxRevenueRecognitionTytm = deptUtils.singleDeptCredit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zjzxRevenueRecognitionTy   = deptUtils.singleDeptCredit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zjzxRevenueRecognitionLytm = deptUtils.singleDeptCredit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zjzxRevenueRecognitionLy   = deptUtils.singleDeptCredit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double zjzxInvoicingRevenueTytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zjzxInvoicingRevenueTy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zjzxInvoicingRevenueLytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zjzxInvoicingRevenueLy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

        //主营业务成本
        double zjzxMBCTytm= deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxMBCTy  = deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxMBCLytm= deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zjzxMBCLy  = deptUtils.singleDeptDebit(Constant.ZJZX,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);

        //毛利率
        double zjzxGrossMarginTytm=(zjzxRevenueRecognitionTytm-zjzxMBCTytm)/zjzxRevenueRecognitionTytm*100;
        double zjzxGrossMarginTy  =(zjzxRevenueRecognitionTy  -zjzxMBCTy  )/zjzxRevenueRecognitionTy  *100;
        double zjzxGrossMarginLytm=(zjzxRevenueRecognitionLytm-zjzxMBCLytm)/zjzxRevenueRecognitionLytm*100;
        double zjzxGrossMarginLy  =(zjzxRevenueRecognitionLy  -zjzxMBCLy  )/zjzxRevenueRecognitionLy  *100;

        deptUtils.outputMB1Excel(Constant.ZJZX_NAME,workbook,sheet,24,
                zjzxRevenueRecognitionTytm,
                zjzxRevenueRecognitionTy,
                zjzxRevenueRecognitionLytm,
                zjzxRevenueRecognitionLy,
                zjzxInvoicingRevenueTytm,
                zjzxInvoicingRevenueTy,
                zjzxInvoicingRevenueLytm,
                zjzxInvoicingRevenueLy,
                zjzxMBCTytm,
                zjzxMBCTy,
                zjzxMBCLytm,
                zjzxMBCLy,
                zjzxGrossMarginTytm,
                zjzxGrossMarginTy,
                zjzxGrossMarginLytm,
                zjzxGrossMarginLy);
    }

    private void getZbdl(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double zbdlRevenueRecognitionTytm = deptUtils.singleDeptCredit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zbdlRevenueRecognitionTy   = deptUtils.singleDeptCredit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zbdlRevenueRecognitionLytm = deptUtils.singleDeptCredit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);
        double zbdlRevenueRecognitionLy   = deptUtils.singleDeptCredit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double zbdlInvoicingRevenueTytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoicingRevenueTy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoicingRevenueLytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double zbdlInvoicingRevenueLy   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

        //主营业务成本
        double zbdlMBCTytm= deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlMBCTy  = deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlMBCLytm= deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);
        double zbdlMBCLy  = deptUtils.singleDeptDebit(Constant.ZBDL,connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), Constant.MAIN_BUSINESS_COST);

        //毛利率
        double zbdlGrossMarginTytm=(zbdlRevenueRecognitionTytm-zbdlMBCTytm)/zbdlRevenueRecognitionTytm*100;
        double zbdlGrossMarginTy  =(zbdlRevenueRecognitionTy  -zbdlMBCTy  )/zbdlRevenueRecognitionTy  *100;
        double zbdlGrossMarginLytm=(zbdlRevenueRecognitionLytm-zbdlMBCLytm)/zbdlRevenueRecognitionLytm*100;
        double zbdlGrossMarginLy  =(zbdlRevenueRecognitionLy  -zbdlMBCLy  )/zbdlRevenueRecognitionLy  *100;
        deptUtils.outputMB1Excel(Constant.ZBDL_NAME,workbook,sheet,23,
                zbdlRevenueRecognitionTytm,
                zbdlRevenueRecognitionTy,
                zbdlRevenueRecognitionLytm,
                zbdlRevenueRecognitionLy,
                zbdlInvoicingRevenueTytm,
                zbdlInvoicingRevenueTy,
                zbdlInvoicingRevenueLytm,
                zbdlInvoicingRevenueLy,
                zbdlMBCTytm,
                zbdlMBCTy,
                zbdlMBCLytm,
                zbdlMBCLy,
                zbdlGrossMarginTytm,
                zbdlGrossMarginTy,
                zbdlGrossMarginLytm,
                zbdlGrossMarginLy);
    }

    private void getSljlMBC(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] sljlRevenueRecognitionTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME);
        double[] sljlRevenueRecognitionTy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME);
        double[] sljlRevenueRecognitionLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME);
        double[] sljlRevenueRecognitionLy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME);

        //挂账收入
        double[] sljlInvoicingRevenueTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] sljlInvoicingRevenueTy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] sljlInvoicingRevenueLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);
        double[] sljlInvoicingRevenueLy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.MAIN_BUSINESS_INCOME,Constant.MAIN_BUSINESS_INCOME_INVOICING);

        //主营业务成本
        double[] sljlMBCTytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCTy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCLytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);
        double[] sljlMBCLy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), Constant.MAIN_BUSINESS_COST);

        //毛利率
        double[] sljlGrossMarginTytm=new double[sljlMBCTytm.length];
        double[] sljlGrossMarginTy  =new double[sljlMBCTytm.length];
        double[] sljlGrossMarginLytm=new double[sljlMBCTytm.length];
        double[] sljlGrossMarginLy  =new double[sljlMBCTytm.length];

        for (int i = 0; i < sljlMBCTytm.length; i++) {
            sljlGrossMarginTytm[i]=(sljlRevenueRecognitionTytm[i]-sljlMBCTytm[i])/sljlRevenueRecognitionTytm[i]*100;
            sljlGrossMarginTy  [i]=(sljlRevenueRecognitionTy  [i]-sljlMBCTy  [i])/sljlRevenueRecognitionTy  [i]*100;
            sljlGrossMarginLytm[i]=(sljlRevenueRecognitionLytm[i]-sljlMBCLytm[i])/sljlRevenueRecognitionLytm[i]*100;
            sljlGrossMarginLy  [i]=(sljlRevenueRecognitionLy  [i]-sljlMBCLy  [i])/sljlRevenueRecognitionLy  [i]*100;
        }

        deptUtils.outputMB1Excel(Constant.SLJL_BRANCH_OFFICE_NAME,workbook,sheet,13,
                sljlRevenueRecognitionTytm,
                sljlRevenueRecognitionTy,
                sljlRevenueRecognitionLytm,
                sljlRevenueRecognitionLy  ,
                sljlInvoicingRevenueTytm,
                sljlInvoicingRevenueTy,
                sljlInvoicingRevenueLytm,
                sljlInvoicingRevenueLy  ,
                sljlMBCTytm,
                sljlMBCTy,
                sljlMBCLytm,
                sljlMBCLy  ,
                sljlGrossMarginTytm,
                sljlGrossMarginTy,
                sljlGrossMarginLytm,
                sljlGrossMarginLy  );
    }

    private void getSljlOverhead(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
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

        deptUtils.outputMBExel(Constant.SLJL_MANAGE_DEPT_NAME,workbook,sheet,6,14,
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