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

    //评估咨询
    private void getPgzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] pgzxRevenueRecognitionTytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001");
        double[] pgzxRevenueRecognitionTy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001");
        double[] pgzxRevenueRecognitionLytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001");
        double[] pgzxRevenueRecognitionLy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001");

        //挂账收入
        double[] pgzxInvoicingRevenueTytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001","600101");
        double[] pgzxInvoicingRevenueTy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001","600101");
        double[] pgzxInvoicingRevenueLytm = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001","600101");
        double[] pgzxInvoicingRevenueLy   = pgzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),"6001","600101");

        //主营业务成本
        double[] pgzxMBCTytm= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxMBCTy  = pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxMBCLytm= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");
        double[] pgzxMBCLy  = pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3), "6401","1403");

        //管理费用（设计）
        double pgzxOCTytm= pgzxOCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "6602");
        double pgzxOCTy  = pgzxOCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "6602");
        double pgzxOCLytm= pgzxOCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "6602");
        double pgzxOCLy  = pgzxOCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "6602");

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
    private double pgzxOCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook,String ledgerAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double[] pgzxMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                  String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount,String ledgerAccount1) throws SQLException {
        return deptUtils.pgzxDeptSum1(connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1, ledgerAccount1));
    }
    private double[] pgzxMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1, ledgerAccount));
    }
    private double[] pgzxMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount,String specificAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    //石大设计
    private void getSdsj(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double sdsjRevenueRecognitionTytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"6001");
        double sdsjRevenueRecognitionTy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"6001");
        double sdsjRevenueRecognitionLytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"6001");
        double sdsjRevenueRecognitionLy   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"6001");

        //挂账收入
        double sdsjInvoicingRevenueTytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoicingRevenueTy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoicingRevenueLytm = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");
        double sdsjInvoicingRevenueLy   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),"1122","112201");

        //存货
        double sdsjMBCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjMBCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjMBCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");
        double sdsjMBCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403");

        //管理费用
        double sdsjOCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602");
        double sdsjOCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602");
        double sdsjOCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602");
        double sdsjOCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602");

        //销售费用
        double sdsjSCTytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601");
        double sdsjSCTy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601");
        double sdsjSCLytm= deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601");
        double sdsjSCLy  = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601");

        double sdsjDevelopTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "5301");
        double sdsjDevelopTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "5301");
        double sdsjDevelopLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "5301");
        double sdsjDevelopLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "5301");


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
                sdsjMBCTytm,
                sdsjMBCTy,
                sdsjMBCLytm,
                sdsjMBCLy,
                sdsjGrossMarginTytm,
                sdsjGrossMarginTy,
                sdsjGrossMarginLytm,
                sdsjGrossMarginLy);
    }

    //恒远检测
    private void getHyjc(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double hyjcRevenueRecognitionTytm = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"6001");
        double hyjcRevenueRecognitionTy   = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"6001");
        double hyjcRevenueRecognitionLytm = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"6001");
        double hyjcRevenueRecognitionLy   = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"6001");

        //挂账收入
        double hyjcInvoicingRevenueTytm = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"6001","600101");
        double hyjcInvoicingRevenueTy   = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),"6001","600101");
        double hyjcInvoicingRevenueLytm = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"6001","600101");
        double hyjcInvoicingRevenueLy   = singleDeptCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),"6001","600101");

        //主营业务成本
        double hyjcMBCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcMBCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcMBCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401");
        double hyjcMBCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401");

        //管理费用
        double hyjcOCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6602");
        double hyjcOCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6602");
        double hyjcOCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6602");
        double hyjcOCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6602");

        //销售费用
        double hyjcSCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6601");
        double hyjcSCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6601");
        double hyjcSCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6601");
        double hyjcSCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6601");

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
                hyjcMBCTytm,
                hyjcMBCTy,
                hyjcMBCLytm,
                hyjcMBCLy,
                hyjcGrossMarginTytm,
                hyjcGrossMarginTy,
                hyjcGrossMarginLytm,
                hyjcGrossMarginLy);
    }
    private double singleDeptCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                    String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double singleDeptCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                    String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    //华海安科
    private void getHhak(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] hhakRevenueRecognitionTytm= hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), "6001");
        double[] hhakRevenueRecognitionTy  = hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), "6001");
        double[] hhakRevenueRecognitionLytm= hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), "6001");
        double[] hhakRevenueRecognitionLy  = hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), "6001");

        //挂账收入
        double[] hhakInvoicingRevenueTytm= hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), "6001","600101");
        double[] hhakInvoicingRevenueTy  = hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(1), "6001","600101");
        double[] hhakInvoicingRevenueLytm= hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), "6001","600101");
        double[] hhakInvoicingRevenueLy  = hhakCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(1), "6001","600101");

        //主营业务成本
        double[] hhakMBCTytm= hhakMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double[] hhakMBCTy  = hhakMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401");
        double[] hhakMBCLytm= hhakMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");
        double[] hhakMBCLy  = hhakMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401");

        //管理费用
        double hhakOCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6602");
        double hhakOCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6602");
        double hhakOCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6602");
        double hhakOCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6602");

        //销售费用
        double hhakSCTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6601");
        double hhakSCTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6601");
        double hhakSCLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6601");
        double hhakSCLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6601");

        //营业总成本
        hhakMBCTytm[0]+=(hhakOCTytm+hhakSCTytm);
        hhakMBCTy  [0]+=(hhakOCTy  +hhakSCTy  );
        hhakMBCLytm[0]+=(hhakOCLytm+hhakSCLytm);
        hhakMBCLy  [0]+=(hhakOCLy  +hhakSCLy  );

        //北京营业总成本
        hhakMBCTytm[1]+=hhakSCTytm;
        hhakMBCTy  [1]+=hhakSCTy  ;
        hhakMBCLytm[1]+=hhakSCLytm;
        hhakMBCLy  [1]+=hhakSCLy  ;

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
    private double[] hhakMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                  String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.hhakDeptSum(connection,commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount));
    }
    private double[] hhakCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.hhakDeptSum(connection,commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount));
    }
    private double[] hhakCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.hhakDeptSum(connection,commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    //造价咨询
    private void getZjzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double zjzxRevenueRecognitionTytm = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zjzxRevenueRecognitionTy   = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zjzxRevenueRecognitionLytm = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zjzxRevenueRecognitionLy   = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001");

        //挂账收入
        double zjzxInvoicingRevenueTytm = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zjzxInvoicingRevenueTy   = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zjzxInvoicingRevenueLytm = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zjzxInvoicingRevenueLy   = zjzxMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001","600101");

        //主营业务成本
        double zjzxMBCTytm= zjzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxMBCTy  = zjzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxMBCLytm= zjzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zjzxMBCLy  = zjzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");

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
    private double zjzxMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double zjzxMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double zjzxMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));

    }

    //招标代理
    private void getZbdl(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double zbdlRevenueRecognitionTytm = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zbdlRevenueRecognitionTy   = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zbdlRevenueRecognitionLytm = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001");
        double zbdlRevenueRecognitionLy   = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001");

        //挂账收入
        double zbdlInvoicingRevenueTytm = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zbdlInvoicingRevenueTy   = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zbdlInvoicingRevenueLytm = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001","600101");
        double zbdlInvoicingRevenueLy   = zbdlMBCCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),"6001","600101");

        //主营业务成本
        double zbdlMBCTytm= zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlMBCTy  = zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlMBCLytm= zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");
        double zbdlMBCLy  = zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401");

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
    private double zbdlMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double zbdlMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount));
    }
    private double zbdlMBCCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));

    }

    /**
     * 监理-生产部门
     */
    private void getSljlMBC(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //确认收入
        double[] sljlRevenueRecognitionTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001");
        double[] sljlRevenueRecognitionTy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001");
        double[] sljlRevenueRecognitionLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001");
        double[] sljlRevenueRecognitionLy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001");

        //挂账收入
        double[] sljlInvoicingRevenueTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001","600101");
        double[] sljlInvoicingRevenueTy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001","600101");
        double[] sljlInvoicingRevenueLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001","600101");
        double[] sljlInvoicingRevenueLy   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),"6001","600101");

        //主营业务成本
        double[] sljlMBCTytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCTy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCLytm= deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");
        double[] sljlMBCLy  = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401");

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

    /**
     * 监理-管理部门
     */
    private void getSljlOverhead(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //监理管理费用
        double[] sljlOverheadTytm= deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadTy  = deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                           Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadLytm= deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), "6602");
        double[] sljlOverheadLy  = deptUtils.sljlOverheadDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),  Constant.ACCOUNTING_BOOK.get(0), "6602");

        //监理销售费用
        double sljlSellingTytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingTy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingLytm= deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601");
        double sljlSellingLy  = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601");

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