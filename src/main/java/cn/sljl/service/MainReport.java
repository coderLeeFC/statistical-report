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
 * @author wangeqiu
 * @version 1.0
 * @date 2024/3/17 11:42
 */
public class MainReport {
    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();
        DateUtils dateUtils = new DateUtils();

        titleModel.createCFSTitle(workbook,sheet,20,Constant.COMPANY_CFS,Constant.TITLE_SIXTH_RIGHT_MAIN);

        getSljl(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZbdl(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getZjzx(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHhak(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getHyjc(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getSdsj(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
        getPgzx(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    }

    private void getPgzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double[] pgzxInTytm = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] pgzxInLytm = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double[] beginOfLastYear = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfThisYear = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double[] beginOfThisYearThisMonthDebit = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfThisYearThisMonthDebit   = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfLastYearThisMonthDebit = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfLastYearThisMonthDebit   = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double[] beginOfThisYearThisMonthCredit = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfThisYearThisMonthCredit   = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfLastYearThisMonthCredit = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfLastYearThisMonthCredit   = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double[] beginOfThisYearThisMonth=new double[beginOfLastYear.length];
        double[] endOfThisYearThisMonth  =new double[beginOfLastYear.length];
        double[] beginOfLastYearThisMonth=new double[beginOfLastYear.length];
        double[] endOfLastYearThisMonth  =new double[beginOfLastYear.length];

        for (int i = 0; i < beginOfLastYear.length; i++) {
            beginOfThisYearThisMonth[i] = beginOfThisYear[i]+ beginOfThisYearThisMonthDebit[i] -beginOfThisYearThisMonthCredit[i];
            endOfThisYearThisMonth  [i] = beginOfThisYear[i]+ endOfThisYearThisMonthDebit  [i] -endOfThisYearThisMonthCredit  [i];
            beginOfLastYearThisMonth[i]= beginOfLastYear [i]+ beginOfLastYearThisMonthDebit[i] -beginOfLastYearThisMonthCredit[i];
            endOfLastYearThisMonth  [i]= beginOfLastYear [i]+ endOfLastYearThisMonthDebit  [i] -endOfLastYearThisMonthCredit  [i];
        }

        //期末合同资产
        //1.应收账款-暂估（期初）
        double[] beginOfLastYear1 = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfThisYear1 = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double[] beginOfThisYearThisMonthDebit1 = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfThisYearThisMonthDebit1   = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfLastYearThisMonthDebit1 = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfLastYearThisMonthDebit1   = pgzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double[] beginOfThisYearThisMonthCredit1 = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfThisYearThisMonthCredit1   = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfLastYearThisMonthCredit1 = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfLastYearThisMonthCredit1   = pgzxCredit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double[] beginOfThisYearThisMonth1=new double[beginOfLastYear.length];
        double[] endOfThisYearThisMonth1  =new double[beginOfLastYear.length];
        double[] beginOfLastYearThisMonth1=new double[beginOfLastYear.length];
        double[] endOfLastYearThisMonth1  =new double[beginOfLastYear.length];

        for (int i = 0; i < beginOfLastYear.length; i++) {
            beginOfThisYearThisMonth1[i]= beginOfThisYear1[i]+ beginOfThisYearThisMonthDebit1[i] -beginOfThisYearThisMonthCredit1[i];
            endOfThisYearThisMonth1  [i]= beginOfThisYear1[i]+ endOfThisYearThisMonthDebit1  [i] -endOfThisYearThisMonthCredit1  [i];
            beginOfLastYearThisMonth1[i]= beginOfLastYear1[i]+ beginOfLastYearThisMonthDebit1[i] -beginOfLastYearThisMonthCredit1[i];
            endOfLastYearThisMonth1  [i]= beginOfLastYear1[i]+ endOfLastYearThisMonthDebit1  [i] -endOfLastYearThisMonthCredit1  [i];
        }

        deptUtils.outputMBI(Constant.PGZX,workbook,sheet,22,
                pgzxInTytm,
                endOfThisYearThisMonthCredit,
                pgzxInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1);
    }
    private double[] pgzxCredit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                   String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }
    private double[] pgzxDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return deptUtils.pgzxDeptSum(connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    private void getZjzx(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zjzxInTytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zjzxInLytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double beginOfLastYear = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfThisYear = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double beginOfThisYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double beginOfThisYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double beginOfThisYearThisMonth = beginOfThisYear+ beginOfThisYearThisMonthDebit -beginOfThisYearThisMonthCredit;
        double endOfThisYearThisMonth   = beginOfThisYear+ endOfThisYearThisMonthDebit   -endOfThisYearThisMonthCredit  ;
        double beginOfLastYearThisMonth= beginOfLastYear + beginOfLastYearThisMonthDebit -beginOfLastYearThisMonthCredit;
        double endOfLastYearThisMonth  = beginOfLastYear + endOfLastYearThisMonthDebit   -endOfLastYearThisMonthCredit  ;

        //期末合同资产
        //1.应收账款-暂估（期初）
        double beginOfLastYear1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfThisYear1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double beginOfThisYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double beginOfThisYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZJZX, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double beginOfThisYearThisMonth1= beginOfThisYear1+ beginOfThisYearThisMonthDebit1 -beginOfThisYearThisMonthCredit1;
        double endOfThisYearThisMonth1  = beginOfThisYear1+ endOfThisYearThisMonthDebit1   -endOfThisYearThisMonthCredit1  ;
        double beginOfLastYearThisMonth1= beginOfLastYear1+ beginOfLastYearThisMonthDebit1 -beginOfLastYearThisMonthCredit1;
        double endOfLastYearThisMonth1  = beginOfLastYear1+ endOfLastYearThisMonthDebit1   -endOfLastYearThisMonthCredit1  ;

        deptUtils.outputMBI(Constant.ZJZX_NAME,workbook,sheet,17,
                zjzxInTytm,
                endOfThisYearThisMonthCredit,
                zjzxInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1  ,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1  );
    }

    private void getZbdl(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double zbdlInTytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double zbdlInLytm = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double beginOfLastYear = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfThisYear = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double beginOfThisYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double beginOfThisYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double beginOfThisYearThisMonth = beginOfThisYear+ beginOfThisYearThisMonthDebit -beginOfThisYearThisMonthCredit;
        double endOfThisYearThisMonth   = beginOfThisYear+ endOfThisYearThisMonthDebit   -endOfThisYearThisMonthCredit  ;
        double beginOfLastYearThisMonth= beginOfLastYear + beginOfLastYearThisMonthDebit -beginOfLastYearThisMonthCredit;
        double endOfLastYearThisMonth  = beginOfLastYear + endOfLastYearThisMonthDebit   -endOfLastYearThisMonthCredit  ;

        //期末合同资产
        //1.应收账款-暂估（期初）
        double beginOfLastYear1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfThisYear1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double beginOfThisYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double beginOfThisYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql,Constant.ZBDL, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double beginOfThisYearThisMonth1= beginOfThisYear1+ beginOfThisYearThisMonthDebit1 -beginOfThisYearThisMonthCredit1;
        double endOfThisYearThisMonth1  = beginOfThisYear1+ endOfThisYearThisMonthDebit1   -endOfThisYearThisMonthCredit1  ;
        double beginOfLastYearThisMonth1= beginOfLastYear1+ beginOfLastYearThisMonthDebit1 -beginOfLastYearThisMonthCredit1;
        double endOfLastYearThisMonth1  = beginOfLastYear1+ endOfLastYearThisMonthDebit1   -endOfLastYearThisMonthCredit1  ;

        deptUtils.outputMBI(Constant.ZBDL_NAME,workbook,sheet,16,
                zbdlInTytm,
                endOfThisYearThisMonthCredit,
                zbdlInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1  ,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1  );



    }

    private void getSdsj(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double sdsjInTytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double sdsjInLytm = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double beginOfLastYear = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfThisYear = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double beginOfThisYearThisMonthDebit = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthDebit   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthDebit = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthDebit   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double beginOfThisYearThisMonthCredit = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthCredit   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthCredit = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthCredit   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double beginOfThisYearThisMonth = beginOfThisYear+ beginOfThisYearThisMonthDebit -beginOfThisYearThisMonthCredit;
        double endOfThisYearThisMonth   = beginOfThisYear+ endOfThisYearThisMonthDebit   -endOfThisYearThisMonthCredit  ;
        double beginOfLastYearThisMonth= beginOfLastYear + beginOfLastYearThisMonthDebit -beginOfLastYearThisMonthCredit;
        double endOfLastYearThisMonth  = beginOfLastYear + endOfLastYearThisMonthDebit   -endOfLastYearThisMonthCredit  ;

        //期末合同资产
        //1.应收账款-暂估（期初）
        double beginOfLastYear1 = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfThisYear1 = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double beginOfThisYearThisMonthDebit1 = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthDebit1   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthDebit1 = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthDebit1   = deptUtils.singleDeptDebit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double beginOfThisYearThisMonthCredit1 = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthCredit1   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthCredit1 = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthCredit1   = deptUtils.singleDeptCredit(Constant.SDSJ_DEPT_INCOME.get(4),connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double beginOfThisYearThisMonth1= beginOfThisYear1+ beginOfThisYearThisMonthDebit1 -beginOfThisYearThisMonthCredit1;
        double endOfThisYearThisMonth1  = beginOfThisYear1+ endOfThisYearThisMonthDebit1   -endOfThisYearThisMonthCredit1  ;
        double beginOfLastYearThisMonth1= beginOfLastYear1+ beginOfLastYearThisMonthDebit1 -beginOfLastYearThisMonthCredit1;
        double endOfLastYearThisMonth1  = beginOfLastYear1+ endOfLastYearThisMonthDebit1   -endOfLastYearThisMonthCredit1  ;

        deptUtils.outputMBI(Constant.SDSJ,workbook,sheet,21,
                sdsjInTytm,
                endOfThisYearThisMonthCredit,
                sdsjInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1  ,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1  );
    }

    private void getHyjc(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hyjcInTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hyjcInLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double beginOfLastYear = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfThisYear = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double beginOfThisYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double beginOfThisYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double beginOfThisYearThisMonth = beginOfThisYear+ beginOfThisYearThisMonthDebit -beginOfThisYearThisMonthCredit;
        double endOfThisYearThisMonth   = beginOfThisYear+ endOfThisYearThisMonthDebit   -endOfThisYearThisMonthCredit  ;
        double beginOfLastYearThisMonth= beginOfLastYear + beginOfLastYearThisMonthDebit -beginOfLastYearThisMonthCredit;
        double endOfLastYearThisMonth  = beginOfLastYear + endOfLastYearThisMonthDebit   -endOfLastYearThisMonthCredit  ;

        //期末合同资产
        //1.应收账款-暂估（期初）
        double beginOfLastYear1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfThisYear1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double beginOfThisYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double beginOfThisYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(2),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double beginOfThisYearThisMonth1= beginOfThisYear1+ beginOfThisYearThisMonthDebit1 -beginOfThisYearThisMonthCredit1;
        double endOfThisYearThisMonth1  = beginOfThisYear1+ endOfThisYearThisMonthDebit1   -endOfThisYearThisMonthCredit1  ;
        double beginOfLastYearThisMonth1= beginOfLastYear1+ beginOfLastYearThisMonthDebit1 -beginOfLastYearThisMonthCredit1;
        double endOfLastYearThisMonth1  = beginOfLastYear1+ endOfLastYearThisMonthDebit1   -endOfLastYearThisMonthCredit1  ;

        deptUtils.outputMBI(Constant.HYJC,workbook,sheet,20,
                hyjcInTytm,
                endOfThisYearThisMonthCredit,
                hyjcInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1  ,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1  );
    }

    private void getHhak(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double hhakInTytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double hhakInLytm = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double beginOfLastYear = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfThisYear = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double beginOfThisYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthDebit = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthDebit   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double beginOfThisYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfThisYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double beginOfLastYearThisMonthCredit = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double endOfLastYearThisMonthCredit   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double beginOfThisYearThisMonth = beginOfThisYear+ beginOfThisYearThisMonthDebit -beginOfThisYearThisMonthCredit;
        double endOfThisYearThisMonth   = beginOfThisYear+ endOfThisYearThisMonthDebit   -endOfThisYearThisMonthCredit  ;
        double beginOfLastYearThisMonth= beginOfLastYear + beginOfLastYearThisMonthDebit -beginOfLastYearThisMonthCredit;
        double endOfLastYearThisMonth  = beginOfLastYear + endOfLastYearThisMonthDebit   -endOfLastYearThisMonthCredit  ;

        //期末合同资产
        //1.应收账款-暂估（期初）
        double beginOfLastYear1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfThisYear1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double beginOfThisYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthDebit1 = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthDebit1   = deptUtils.singleDeptDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double beginOfThisYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfThisYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double beginOfLastYearThisMonthCredit1 = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double endOfLastYearThisMonthCredit1   = deptUtils.singleDeptCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(1),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double beginOfThisYearThisMonth1= beginOfThisYear1+ beginOfThisYearThisMonthDebit1 -beginOfThisYearThisMonthCredit1;
        double endOfThisYearThisMonth1  = beginOfThisYear1+ endOfThisYearThisMonthDebit1   -endOfThisYearThisMonthCredit1  ;
        double beginOfLastYearThisMonth1= beginOfLastYear1+ beginOfLastYearThisMonthDebit1 -beginOfLastYearThisMonthCredit1;
        double endOfLastYearThisMonth1  = beginOfLastYear1+ endOfLastYearThisMonthDebit1   -endOfLastYearThisMonthCredit1  ;

        deptUtils.outputMBI(Constant.HHAK_DEPT_NAME.get(0),workbook,sheet,18,
                hhakInTytm,
                endOfThisYearThisMonthCredit,
                hhakInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1  ,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1  );
    }

    private void getSljl(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //现金流入：应收账款-开票（贷方）
        double[] sljlMBCInTytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInLytm = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //期末应收账款
        //1.应收账款-开票（期初）
        double[] beginOfLastYear = deptUtils.sljlManufactureDebit(connection, commonSql,  dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfThisYear = deptUtils.sljlManufactureDebit(connection, commonSql,  dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //2.应收账款-开票（借方累计）
        double[] beginOfThisYearThisMonthDebit = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfThisYearThisMonthDebit   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfLastYearThisMonthDebit = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfLastYearThisMonthDebit   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //3.应收账款-开票（贷方累计）
        double[] beginOfThisYearThisMonthCredit = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfThisYearThisMonthCredit   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] beginOfLastYearThisMonthCredit = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] endOfLastYearThisMonthCredit   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        //4.期末应收账款
        double[] beginOfThisYearThisMonth=new double[beginOfLastYear.length];
        double[] endOfThisYearThisMonth  =new double[beginOfLastYear.length];
        double[] beginOfLastYearThisMonth=new double[beginOfLastYear.length];
        double[] endOfLastYearThisMonth  =new double[beginOfLastYear.length];

        for (int i = 0; i < beginOfLastYear.length; i++) {
            beginOfThisYearThisMonth[i] = beginOfThisYear[i]+ beginOfThisYearThisMonthDebit[i] -beginOfThisYearThisMonthCredit[i];
            endOfThisYearThisMonth  [i] = beginOfThisYear[i]+ endOfThisYearThisMonthDebit  [i] -endOfThisYearThisMonthCredit  [i];
            beginOfLastYearThisMonth[i]= beginOfLastYear [i]+ beginOfLastYearThisMonthDebit[i] -beginOfLastYearThisMonthCredit[i];
            endOfLastYearThisMonth  [i]= beginOfLastYear [i]+ endOfLastYearThisMonthDebit  [i] -endOfLastYearThisMonthCredit  [i];
        }

        //期末合同资产
        //1.应收账款-暂估（期初）
        double[] beginOfLastYear1 = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END), dateUtils.getStartOfLastYear(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfThisYear1 = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getStartOfYear(Constant.THIS_MONTH_END),     dateUtils.getStartOfYear(Constant.THIS_MONTH_END),    Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //2.应收账款-暂估（借方累计）
        double[] beginOfThisYearThisMonthDebit1 = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfThisYearThisMonthDebit1   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfLastYearThisMonthDebit1 = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfLastYearThisMonthDebit1   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //3.应收账款-暂估（贷方累计）
        double[] beginOfThisYearThisMonthCredit1 = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , dateUtils.getLastMonthEnd(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfThisYearThisMonthCredit1   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END) , Constant.THIS_MONTH_END,                            Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] beginOfLastYearThisMonthCredit1 = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getLastMonthEnd1(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);
        double[] endOfLastYearThisMonthCredit1   = deptUtils.sljlManufactureCredit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),   Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_ESTIMATED);

        //4.期末合同资产
        double[] beginOfThisYearThisMonth1=new double[beginOfLastYear.length];
        double[] endOfThisYearThisMonth1  =new double[beginOfLastYear.length];
        double[] beginOfLastYearThisMonth1=new double[beginOfLastYear.length];
        double[] endOfLastYearThisMonth1  =new double[beginOfLastYear.length];

        for (int i = 0; i < beginOfLastYear.length; i++) {
            beginOfThisYearThisMonth1[i]= beginOfThisYear1[i]+ beginOfThisYearThisMonthDebit1[i] -beginOfThisYearThisMonthCredit1[i];
            endOfThisYearThisMonth1  [i]= beginOfThisYear1[i]+ endOfThisYearThisMonthDebit1  [i] -endOfThisYearThisMonthCredit1  [i];
            beginOfLastYearThisMonth1[i]= beginOfLastYear1[i]+ beginOfLastYearThisMonthDebit1[i] -beginOfLastYearThisMonthCredit1[i];
            endOfLastYearThisMonth1  [i]= beginOfLastYear1[i]+ endOfLastYearThisMonthDebit1  [i] -endOfLastYearThisMonthCredit1  [i];
        }

        deptUtils.outputMBI(Constant.SLJL_BRANCH_OFFICE_NAME,workbook,sheet,6,
                sljlMBCInTytm,
                endOfThisYearThisMonthCredit,
                sljlMBCInLytm,
                endOfLastYearThisMonthCredit  ,
                beginOfThisYearThisMonth,
                endOfThisYearThisMonth  ,
                beginOfLastYearThisMonth,
                endOfLastYearThisMonth  ,
                beginOfThisYearThisMonth1,
                endOfThisYearThisMonth1,
                beginOfLastYearThisMonth1,
                endOfLastYearThisMonth1);
    }



}
