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
 * @date 2024/5/6 11:04
 */
public class InvoiceReport {
    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        DateUtils dateUtils = new DateUtils();
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();

        titleModel.createSalaryTitle(workbook, sheet, 8, "生产经营部综合统计报表");

        double[] sljlMBCInTytm = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfQuarter(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInTy   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                         Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInLytm = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfQuarter1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);
        double[] sljlMBCInLy   = deptUtils.sljlManufactureDebit(connection, commonSql, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END),Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5),Constant.ACCOUNTS_RECEIVABLE,Constant.ACCOUNTS_RECEIVABLE_INVOICING);

        deptUtils.outputSome(Constant.SLJL_BRANCH_OFFICE_NAME,workbook,sheet,5,
                sljlMBCInTytm,
                sljlMBCInTy  ,
                sljlMBCInLytm,
                sljlMBCInLy  );
    }
}
