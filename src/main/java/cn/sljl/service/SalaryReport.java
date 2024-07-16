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
import java.util.Arrays;
import java.util.List;

/**
 * 营业成本-工资性支出
 * @author wangeqiu
 * @version 1.0
 * @date 2024/4/29 08:48
 */
public class SalaryReport {
    public void mainMethod(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, TitleModel titleModel) throws SQLException {
        CommonSql commonSql = new CommonSql();
        DeptUtils deptUtils = new DeptUtils();
        DateUtils dateUtils = new DateUtils();

    titleModel.createSalaryTitle(workbook, sheet, 8,Constant.SALARY_TITLE_RIGHT);

    getSljlOCSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getSLJLSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getZbdlSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getZjzxSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getHHAKSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getHYJCSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getSDSJSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);
    getPgzxSalary(connection,workbook,sheet,dateUtils,commonSql, deptUtils);

    }

    private void getPgzxSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        double[] pgzxTytm =new double[3];
        double[] pgzxTy   =new double[3];
        double[] pgzxLytm =new double[3];
        double[] pgzxLy   =new double[3];
        //主营业务成本
        //监理
        pgzxTytm[1]= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        pgzxTy  [1]= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        pgzxLytm[1]= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        pgzxLy  [1]= pgzxMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));

        //设计
        pgzxTytm[2]= pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "1403",Arrays.asList("14030124","14030125","14030126","14030127"));
        pgzxTy  [2]= pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "1403",Arrays.asList("14030124","14030125","14030126","14030127"));
        pgzxLytm[2]= pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "1403",Arrays.asList("14030124","14030125","14030126","14030127"));
        pgzxLy  [2]= pgzxMBCDebit1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "1403",Arrays.asList("14030124","14030125","14030126","14030127"));

        //管理费用-设计
        double pgzxOverheadTytm= pgzxMBCDebit2(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "6602",Arrays.asList("660222","660223","660224","660225"));
        double pgzxOverheadTy  = pgzxMBCDebit2(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "6602",Arrays.asList("660222","660223","660224","660225"));
        double pgzxOverheadLytm= pgzxMBCDebit2(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "6602",Arrays.asList("660222","660223","660224","660225"));
        double pgzxOverheadLy  = pgzxMBCDebit2(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "6602",Arrays.asList("660222","660223","660224","660225"));

        pgzxTytm[2]+=pgzxOverheadTytm;
        pgzxTy  [2]+=pgzxOverheadTy  ;
        pgzxLytm[2]+=pgzxOverheadLytm;
        pgzxLy  [2]+=pgzxOverheadLy  ;

        pgzxTytm[0]=pgzxTytm[1]+pgzxTytm[2];
        pgzxTy  [0]=pgzxTy  [1]+pgzxTy  [2];
        pgzxLytm[0]=pgzxLytm[1]+pgzxLytm[2];
        pgzxLy  [0]=pgzxLy  [1]+pgzxLy  [2];

        deptUtils.outputSome(Constant.PGZX,workbook,sheet,32,
                pgzxTytm,
                pgzxTy,
                pgzxLytm,
                pgzxLy  );

    }
    private double pgzxMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                  String startDate, String endDate, String accountBook, String ledgerAccount,List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.PGZX_DEPT,commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }
    private double pgzxMBCDebit1(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount,List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.SDSJ_DEPT.get(4),commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }
    private double pgzxMBCDebit2(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount,List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.SDSJ_OVERHEAD_DEPT.get(4),commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    private void getZjzxSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        double zjzxTytmSalary= zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zjzxTySalary  = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zjzxLytmSalary= zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zjzxLySalary  = zjzxDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));

        deptUtils.outputSingle(Constant.ZJZX_NAME,workbook,sheet,23,2,
                zjzxTytmSalary,
                zjzxTySalary,
                zjzxLytmSalary,
                zjzxLySalary  );
    }
    private double zjzxDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                             String startDate, String endDate, String accountBook, String ledgerAccount,List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZJZX,commonSql.quadraDebitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    private void getZbdlSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        double zbdlTytmSalary= zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zbdlTySalary  = zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zbdlLytmSalary= zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));
        double zbdlLySalary  = zbdlMBCDebit(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6401",Arrays.asList("640122","640123","640124","640125"));

        deptUtils.outputSingle(Constant.ZBDL_NAME,workbook,sheet,22,2,
                zbdlTytmSalary,
                zbdlTySalary,
                zbdlLytmSalary,
                zbdlLySalary  );
    }
    private double zbdlMBCDebit(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                String startDate, String endDate, String accountBook, String ledgerAccount,List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,Constant.ZBDL,commonSql.quadraDebitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    private void getSDSJSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //管理部门
        double sdsjOverheadTytm=sdsjDeptSumSalary1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602",Arrays.asList("660222","660223","660224"));
        double sdsjOverheadTy  =sdsjDeptSumSalary1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602",Arrays.asList("660222","660223","660224"));
        double sdsjOverheadLytm=sdsjDeptSumSalary1(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602",Arrays.asList("660222","660223","660224"));
        double sdsjOverheadLy  =sdsjDeptSumSalary1(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6602",Arrays.asList("660222","660223","660224"));

        //市场部门
        double sdsjSellingTytm=sdsjDeptSumSalary2(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601",Arrays.asList("660122","660123","660124"));
        double sdsjSellingTy  =sdsjDeptSumSalary2(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601",Arrays.asList("660122","660123","660124"));
        double sdsjSellingLytm=sdsjDeptSumSalary2(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601",Arrays.asList("660122","660123","660124"));
        double sdsjSellingLy  =sdsjDeptSumSalary2(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "6601",Arrays.asList("660122","660123","660124"));

        //生产部门
        double sdsjMBCTytm=sdsjDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403",Arrays.asList("14030124","14030125","14030126"));
        double sdsjMBCTy  =sdsjDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403",Arrays.asList("14030124","14030125","14030126"));
        double sdsjMBCLytm=sdsjDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403",Arrays.asList("14030124","14030125","14030126"));
        double sdsjMBCLy  =sdsjDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3),Constant.ACCOUNTING_BOOK.get(4), "1403",Arrays.asList("14030124","14030125","14030126"));

        //研发费用
        double sdsjDevelopTytm=sdsjDeptSumSalary3(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "5301",Arrays.asList("53010122","53010123","53010124"));
        double sdsjDevelopTy  =sdsjDeptSumSalary3(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(3), "5301",Arrays.asList("53010122","53010123","53010124"));
        double sdsjDevelopLytm=sdsjDeptSumSalary3(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "5301",Arrays.asList("53010122","53010123","53010124"));
        double sdsjDevelopLy  =sdsjDeptSumSalary3(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(3), "5301",Arrays.asList("53010122","53010123","53010124"));


        double sdsjSumTytm=sdsjOverheadTytm+sdsjSellingTytm+sdsjMBCTytm+sdsjDevelopTytm;
        double sdsjSumTy  =sdsjOverheadTy  +sdsjSellingTy  +sdsjMBCTy  +sdsjDevelopTy;
        double sdsjSumLytm=sdsjOverheadLytm+sdsjSellingLytm+sdsjMBCLytm+sdsjDevelopLytm;
        double sdsjSumLy  =sdsjOverheadLy  +sdsjSellingLy  +sdsjMBCLy  +sdsjDevelopLy;

        deptUtils.outputSingle(Constant.SDSJ,workbook,sheet,31,2,
                sdsjSumTytm,
                sdsjSumTy  ,
                sdsjSumLytm,
                sdsjSumLy  );

    }
    //管理部门
    private double sdsjDeptSumSalary1(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum3(connection,
                commonSql.tripleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.tripleDebitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }
    //市场部门
    private double sdsjDeptSumSalary2(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                        String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum4(connection,
                commonSql.tripleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.tripleDebitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }
    //生产部门
    private double sdsjDeptSumSalary(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum2(connection,
                commonSql.tripleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.tripleDebitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    private double sdsjDeptSumSalary3(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                     String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.tripleDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    private void getHYJCSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //管理部门
        double hyjcTytmSalary1= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hyjcTySalary1  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hyjcLytmSalary1= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hyjcLySalary1  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6602",Arrays.asList("660222","660223","660224","660225"));

        //市场部门
        double hyjcTytmSalary2= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hyjcTySalary2  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hyjcLytmSalary2= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hyjcLySalary2  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6601",Arrays.asList("660122","660123","660124","660125"));

        //生产部门
        double hyjcTytmSalary3= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401",Arrays.asList("640122","640123","640124","640125"));
        double hyjcTySalary3  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(2), "6401",Arrays.asList("640122","640123","640124","640125"));
        double hyjcLytmSalary3= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401",Arrays.asList("640122","640123","640124","640125"));
        double hyjcLySalary3  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(2), "6401",Arrays.asList("640122","640123","640124","640125"));

        double hyjcSumTytm=hyjcTytmSalary1+hyjcTytmSalary2+hyjcTytmSalary3;
        double hyjcSumTy  =hyjcTySalary1  +hyjcTySalary2  +hyjcTySalary3  ;
        double hyjcSumLytm=hyjcLytmSalary1+hyjcLytmSalary2+hyjcLytmSalary3;
        double hyjcSumLy  =hyjcLySalary1  +hyjcLySalary2  +hyjcLySalary3  ;

        deptUtils.outputSingle(Constant.HYJC,workbook,sheet,30,2,
                hyjcSumTytm,
                hyjcSumTy  ,
                hyjcSumLytm,
                hyjcSumLy  );
    }

    private void getHHAKSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //管理部门
        double hhakTytmSalary1= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hhakTySalary1  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hhakLytmSalary1= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6602",Arrays.asList("660222","660223","660224","660225"));
        double hhakLySalary1  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6602",Arrays.asList("660222","660223","660224","660225"));

        //市场部门
        double hhakTytmSalary2= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hhakTySalary2  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hhakLytmSalary2= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6601",Arrays.asList("660122","660123","660124","660125"));
        double hhakLySalary2  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6601",Arrays.asList("660122","660123","660124","660125"));

        //生产部门
        double[] hhakTytmSalary3= hhakDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] hhakTySalary3  = hhakDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(1), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] hhakLytmSalary3= hhakDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] hhakLySalary3  = hhakDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(1), "6401",Arrays.asList("640122","640123","640124","640125"));

        hhakTytmSalary3[1]+=hhakTytmSalary2+hhakTytmSalary1;
        hhakTySalary3[1]+=hhakTySalary2  +hhakTySalary1  ;
        hhakLytmSalary3[1]+=hhakLytmSalary2+hhakLytmSalary1;
        hhakLySalary3[1]+=hhakLySalary2  +hhakLySalary1 ;

        hhakTytmSalary3[0]+=hhakTytmSalary2+hhakTytmSalary1;
        hhakTySalary3[0]+=hhakTySalary2    +hhakTySalary1  ;
        hhakLytmSalary3[0]+=hhakLytmSalary2+hhakLytmSalary1;
        hhakLySalary3[0]+=hhakLySalary2    +hhakLySalary1  ;

        deptUtils.outputSalaryExcel(Constant.HHAK_DEPT_NAME,workbook,sheet,24,
                hhakTytmSalary3,
                hhakTySalary3,
                hhakLytmSalary3,
                hhakLySalary3  ,
                hhakTytmSalary1,
                hhakTySalary1,
                hhakLytmSalary1,
                hhakLySalary1  );
    }
    private double[] hhakDeptSumSalary(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.hhakDeptSum(connection,commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    private void getSLJLSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {
        //生产部门
        double[] sljlTytmSalary3=sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] sljlTySalary3  =sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] sljlLytmSalary3=sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401",Arrays.asList("640122","640123","640124","640125"));
        double[] sljlLySalary3  =sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0),Constant.ACCOUNTING_BOOK.get(5), "6401",Arrays.asList("640122","640123","640124","640125"));

        deptUtils.outputSome(Constant.SLJL_BRANCH_OFFICE_NAME,workbook,sheet,12,
                sljlTytmSalary3,
                sljlTySalary3  ,
                sljlLytmSalary3,
                sljlLySalary3  );
    }
    //生产部门
    private double[] sljlDeptSumSalary(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String accountBook1, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.sljlManufactureDeptSum(connection,
                commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.quadraDebitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }
    //管理部门
    private double[] sljlDeptSumSalary(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                       String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.sljlOverheadDeptSum(connection,commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    private void getSljlOCSalary(Connection connection, XSSFWorkbook workbook, XSSFSheet sheet, DateUtils dateUtils, CommonSql commonSql, DeptUtils deptUtils) throws SQLException {

        double[] sljlOverheadTytm=sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6602",Arrays.asList("660222","660223","660224","660225"));
        double[] sljlOverheadTy  =sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6602",Arrays.asList("660222","660223","660224","660225"));
        double[] sljlOverheadLytm=sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6602",Arrays.asList("660222","660223","660224","660225"));
        double[] sljlOverheadLy  =sljlDeptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6602",Arrays.asList("660222","660223","660224","660225"));

        double sljlSellingTytm= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth(Constant.THIS_MONTH_END), Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601",Arrays.asList("660122","660123","660124","660125"));
        double sljlSellingTy  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear(Constant.THIS_MONTH_END),  Constant.THIS_MONTH_END,                          Constant.ACCOUNTING_BOOK.get(0), "6601",Arrays.asList("660122","660123","660124","660125"));
        double sljlSellingLytm= deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfMonth1(Constant.THIS_MONTH_END),dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601",Arrays.asList("660122","660123","660124","660125"));
        double sljlSellingLy  = deptSumSalary(connection, commonSql, deptUtils, dateUtils.getBeginningOfYear1(Constant.THIS_MONTH_END), dateUtils.getEndOfMonth(Constant.THIS_MONTH_END), Constant.ACCOUNTING_BOOK.get(0), "6601",Arrays.asList("660122","660123","660124","660125"));

        sljlOverheadTytm[0]+=sljlSellingTytm;
        sljlOverheadTy  [0]+=sljlSellingTy  ;
        sljlOverheadLytm[0]+=sljlSellingLytm;
        sljlOverheadLy  [0]+=sljlSellingLy  ;

        deptUtils.outputMBExel(Constant.SLJL_MANAGE_DEPT_NAME,workbook,sheet,5,2,
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
     *无部门划分
     */
    private double deptSumSalary(Connection connection, CommonSql commonSql, DeptUtils deptUtils,
                                 String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount) throws SQLException {
        return deptUtils.singleDeptSum(connection,commonSql.quadraDebitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }
}
