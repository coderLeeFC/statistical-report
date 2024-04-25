package cn.sljl.util;

import cn.sljl.mapper.CommonSql;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.sql.Connection;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;

/**
 * 部门求和
 *
 * @author wangeqiu
 * @version 1.0.0
 * @date 2024/04/09 09:58:47
 */
public class DeptUtils {

    /**
     * 主表：监理-生产部门【借方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param accountBook1
     * @param ledgerAccount
     * @param specificAccount
     * @return
     * @throws SQLException
     */
    public double[] sljlManufactureDebit(Connection connection, CommonSql commonSql, String startDate, String endDate,
                                         String accountBook, String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return sljlManufactureDeptSum(connection,
                commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    public double[] sljlManufactureDebit(Connection connection, CommonSql commonSql, String startDate, String endDate,
                                          String accountBook, String accountBook1, String ledgerAccount) throws SQLException {
        return sljlManufactureDeptSum(connection,
                commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount),
                commonSql.debitAmount(startDate,endDate,accountBook1,ledgerAccount));
    }

    /**
     * 主表：监理-生产部门【贷方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param accountBook1
     * @param ledgerAccount
     * @param specificAccount
     * @return
     * @throws SQLException
     */
    public double[] sljlManufactureCredit(Connection connection, CommonSql commonSql, String startDate, String endDate,
                                          String accountBook, String accountBook1, String ledgerAccount, String specificAccount) throws SQLException {
        return sljlManufactureDeptSum(connection,
                commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate,endDate,accountBook1,ledgerAccount,specificAccount));
    }

    public double[] sljlManufactureCredit(Connection connection, CommonSql commonSql,String startDate, String endDate,
                                   String accountBook, String accountBook1, String ledgerAccount) throws SQLException {
        return sljlManufactureDeptSum(connection,
                commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount),
                commonSql.creditAmount(startDate,endDate,accountBook1,ledgerAccount));
    }

    /**
     * 监理-管理部门【借方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     * @throws SQLException
     */
    public double[] sljlOverheadDebit(Connection connection, CommonSql commonSql,String startDate, String endDate,
                                      String accountBook, String ledgerAccount) throws SQLException {
        return sljlOverheadDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    /**
     * 评估咨询【贷方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param accountBook1
     * @param ledgerAccount
     * @param specificAccount
     * @return
     * @throws SQLException
     */
    public double[] pgzxCredit(Connection connection, CommonSql commonSql,String startDate, String endDate,
                               String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return pgzxDeptSum(connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    /**
     * 评估咨询【借方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param accountBook1
     * @param ledgerAccount
     * @param specificAccount
     * @return
     * @throws SQLException
     */
    public double[] pgzxDebit(Connection connection, CommonSql commonSql, String startDate, String endDate,
                              String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return pgzxDeptSum(connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1, ledgerAccount,specificAccount));
    }

    /**
     * 单部门【借方】
     * @param connection
     * @param commonSql
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     * @throws SQLException
     */
    public double singleDeptDebit(Connection connection, CommonSql commonSql, String startDate, String endDate,
                                  String accountBook, String ledgerAccount) throws SQLException {
        return singleDeptSum(connection,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount));
    }

    public double singleDeptDebit(Connection connection, CommonSql commonSql,String startDate,String endDate,
                                  String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(connection,commonSql.debitAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    public double singleDeptDebit(Connection connection, CommonSql commonSql,String deptNum,String startDate, String endDate,
                             String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(connection,deptNum,commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    public double singleDeptCredit(Connection connection, CommonSql commonSql,String deptNum,String startDate, String endDate,
                                   String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(connection,deptNum,commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount));
    }

    public double singleDeptCredit(Connection connection, CommonSql commonSql,String startDate,String endDate,
                                   String accountBook, String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(connection,commonSql.creditAmount(startDate,endDate,accountBook,ledgerAccount,specificAccount));
    }

    public double singleDeptDebit(String deptNum,Connection connection, CommonSql commonSql,String startDate, String endDate,
                                  String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(deptNum,connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1,ledgerAccount,specificAccount));
    }

    public double singleDeptCredit(String deptNum,Connection connection, CommonSql commonSql,String startDate, String endDate,
                                   String accountBook,String accountBook1,String ledgerAccount, String specificAccount) throws SQLException {
        return singleDeptSum(deptNum,connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount,specificAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1,ledgerAccount,specificAccount));
    }

    public double singleDeptDebit(String deptNum,Connection connection, CommonSql commonSql,String startDate, String endDate,
                                  String accountBook,String accountBook1,String ledgerAccount) throws SQLException {
        return singleDeptSum(deptNum,connection,
                commonSql.debitAmount(startDate, endDate, accountBook, ledgerAccount),
                commonSql.debitAmount(startDate, endDate, accountBook1,ledgerAccount));
    }

    public double singleDeptCredit(String deptNum,Connection connection, CommonSql commonSql,String startDate, String endDate,
                                   String accountBook,String accountBook1,String ledgerAccount) throws SQLException {
        return singleDeptSum(deptNum,connection,
                commonSql.creditAmount(startDate, endDate, accountBook, ledgerAccount),
                commonSql.creditAmount(startDate, endDate, accountBook1,ledgerAccount));
    }

    /**
     * pgzx dept sum
     *
     * @param connection 联系
     * @param sql        sql
     * @param sql1       sql1
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:48
     */
    public double[] pgzxDeptSum(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double[] sum = new double[3];
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith("0215")){
                sum[1]+=resultSet.getDouble("amount");
            }
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(Constant.SDSJ_DEPT_INCOME.get(4)))
            sum[2]+=resultSet.getDouble("amount");
        }

        for (int i = 0; i < sum.length; i++) {
            sum[0]+=sum[i];
        }

        for (int i = 0; i < sum.length; i++) {
            sum[i]=sum[i]/10000;
        }

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * pgzx部门sum1
     *
     * @param connection 联系
     * @param sql        sql
     * @param sql1       sql1
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:48
     */
    public double[] pgzxDeptSum1(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double[] sum = new double[3];
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith("0215")){
                sum[1]+=resultSet.getDouble("amount");
            }
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(Constant.SDSJ_DEPT.get(4)))
                sum[2]+=resultSet.getDouble("amount");
        }

        for (int i = 0; i < sum.length; i++) {
            sum[0]+=sum[i];
        }

        for (int i = 0; i < sum.length; i++) {
            sum[i]=sum[i]/10000;
        }

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * pgzx dept sum
     *
     * @param connection 联系
     * @param sql        sql
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:48
     */
    public double pgzxDeptSum(Connection connection, String sql) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(Constant.SDSJ_DEPT_INCOME.get(4))){
                sum+=resultSet.getDouble("amount");
            }
        }
        sum=sum/10000;


        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 监理
     *
     * @param connection 联系
     * @param sql        sql
     * @param sql1       sql1
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:48
     */
    public double[] sljlManufactureDeptSum(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double[] sum = new double[10];
        while (resultSet.next()){
            for (int i = 0; i < 8; i++) {
                if (resultSet.getString("dept_code").startsWith(Constant.SLJL_DEPT_NUM.get(i))){
                    sum[i+1]+=resultSet.getDouble("amount");
                }
            }
        }

        //青岛
        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            sum[9]+=resultSet.getDouble("amount");
        }

        for (int i = 1; i < sum.length; i++) {
            sum[0]+=sum[i];
        }

        //转换为万元
        for (int i = 0; i < sum.length; i++) {
            sum[i]=sum[i]/10000;
        }
        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * sljl间接费用部门总额
     *
     * @param connection 联系
     * @param sql        sql
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double[] sljlOverheadDeptSum(Connection connection,String sql) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double[] sum = new double[6];

        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(0))||
                    resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(1))){
                sum[1]+=resultSet.getDouble("amount");//0101,0102：公司领导，公司办
            }else if (resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(2))){
                sum[2]+=resultSet.getDouble("amount");//0103：生产经营部
            }else if (resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(3))||
                    resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(4))){
                sum[3]+=resultSet.getDouble("amount");//0104,0109：财务资产部，公共业务部
            }else if (resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(5))||
                    resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(6))){
                sum[4]+=resultSet.getDouble("amount");//0105,0108：人力资源部，人才储备部
            }else if (resultSet.getString("dept_code").startsWith(Constant.SLJL_MANAGE_DEPT.get(7))){
                sum[5]+=resultSet.getDouble("amount");//0107：总务管理部
            }
        }

        for (int i = 1; i < 6; i++) {
            sum[0]+=sum[i];
        }

        //转换为万元
        for (int i = 0; i < sum.length; i++) {
            sum[i]=sum[i]/10000;
        }
        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 华海
     *
     * @param connection 联系
     * @param sql        sql
     * @return {@link double[] }
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double[] hhakDeptSum(Connection connection, String sql) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double[] sum = new double[5];
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(Constant.HHAK_DEPT.get(0))||resultSet.getString("dept_code").startsWith(Constant.HHAK_SELLING)){
                sum[1]+=resultSet.getDouble("amount");//北京
            }else if (resultSet.getString("dept_code").startsWith(Constant.HHAK_DEPT.get(1))){
                sum[2]+=resultSet.getDouble("amount");//管道
            }else if (resultSet.getString("dept_code").startsWith(Constant.HHAK_DEPT.get(2))){
                sum[3]+=resultSet.getDouble("amount");//胜利
            }else if (resultSet.getString("dept_code").startsWith(Constant.HHAK_DEPT.get(3))){
                sum[4]+=resultSet.getDouble("amount");//测绘
            }
        }

        for (int i = 1; i < 5; i++) {
            sum[0]+=sum[i];
        }

        for (int i = 0; i < sum.length; i++) {
            sum[i]=sum[i]/10000;
        }
        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }


    /**
     * 恒远
     *
     * @param connection 联系
     * @param sql        sql
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double singleDeptSum(Connection connection, String sql) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 设计收入部门
     *
     * @param connection
     * @param sql
     * @param sql1
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double singleDeptSum(String deptNum,Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            if (!resultSet.getString("dept_code").startsWith(deptNum)){
                sum+=resultSet.getDouble("amount");
            }
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 设计生产部门
     *
     * @param connection
     * @param sql
     * @param sql1
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double singleDeptSum2(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            if (!resultSet.getString("dept_code").startsWith(Constant.SDSJ_DEPT.get(4))){
                sum+=resultSet.getDouble("amount");
            }
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 设计管理部门
     *
     * @param connection
     * @param sql
     * @param sql1
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double singleDeptSum3(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            if (!resultSet.getString("dept_code").startsWith(Constant.SDSJ_OVERHEAD_DEPT.get(4))){
                sum+=resultSet.getDouble("amount");
            }
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 单一部门sum4
     *
     * @param connection 联系
     * @param sql        sql
     * @param sql1       sql1
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:49
     */
    public double singleDeptSum4(Connection connection, String sql, String sql1) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        statement=connection.prepareStatement(sql1);
        resultSet= statement.executeQuery();
        while (resultSet.next()){
            sum+=resultSet.getDouble("amount");
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }

    /**
     * 单部门合计
     *
     * @param connection 联系
     * @param dept       部
     * @param sql        sql
     * @return double
     * @throws SQLException SQLException
     * @author wangeqiu
     * @date 2024/04/09 09:58:50
     */
    public double singleDeptSum(Connection connection,String dept, String sql) throws SQLException {
        PreparedStatement statement = connection.prepareStatement(sql);
        ResultSet resultSet = statement.executeQuery();
        double sum = 0.0;
        while (resultSet.next()){
            if (resultSet.getString("dept_code").startsWith(dept)){
                sum+=resultSet.getDouble("amount");
            }
        }

        sum=sum/10000;

        JdbcUtilsDbcp.release(null,statement,resultSet);
        return sum;
    }


    public void outputMBExel(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,int column,
                             double[] tytm, double[] ty, double[] lytm, double[] ly,
                             double tytm1, double ty1, double lytm1, double ly1){
        TitleModel titleModel = new TitleModel();


        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
//            if (i==0){
//                deptCell.setCellStyle(titleModel.cellCommonStyle(workbook));
//            }else {
//                deptCell.setCellStyle(titleModel.cellCommonStyle3(workbook));
//            }
            deptCell.setCellValue(dept.get(i));

            //数据列
            XSSFCell cell = row.createCell(column);
            if (i==0){
                cell.setCellStyle(titleModel.cellCommonStyle2(workbook));
                cell.setCellValue(tytm[i]);
            }else if (i==dept.size()-1){
                cell.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell.setCellValue(tytm1);
            }else {
                cell.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell.setCellValue(tytm[i]);
            }

            XSSFCell cell1 = row.createCell(column+1);
            if (i==0){
                cell1.setCellStyle(titleModel.cellCommonStyle2(workbook));
                cell1.setCellValue(ty[i]);
            }else if (i==dept.size()-1){
                cell1.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell1.setCellValue(ty1);
            }else {
                cell1.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell1.setCellValue(ty[i]);
            }

            XSSFCell cell2 = row.createCell(column+2);
            if (i==0){
                cell2.setCellStyle(titleModel.cellCommonStyle2(workbook));
                cell2.setCellValue(lytm[i]);
            }else if (i==dept.size()-1){
                cell2.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell2.setCellValue(lytm1);
            }else {
                cell2.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell2.setCellValue(lytm[i]);
            }

            XSSFCell cell3 = row.createCell(column+3);
            if (i==0){
                cell3.setCellStyle(titleModel.cellCommonStyle2(workbook));
                cell3.setCellValue(ly[i]);
            }else if (i==dept.size()-1){
                cell3.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell3.setCellValue(ly1);
            }else {
                cell3.setCellStyle(titleModel.cellCommonStyle3(workbook));
                cell3.setCellValue(ly[i]);
            }

            XSSFCell cell4 = row.createCell(column+4);

            if (i==0){
                cell4.setCellStyle(titleModel.cellCommonStyle2(workbook));
                if (lytm[i]==0){
                    cell4.setCellValue(0);
                }else {
                    cell4.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
                }
            }else if (i==dept.size()-1){
                cell4.setCellStyle(titleModel.cellCommonStyle3(workbook));
                if (lytm1==0){
                    cell4.setCellValue(0);
                }else {
                    cell4.setCellValue((tytm1-lytm1)/lytm1*100);
                }
            }else {
                cell4.setCellStyle(titleModel.cellCommonStyle3(workbook));
                if (lytm[i]==0){
                    cell4.setCellValue(0);
                }else {
                    cell4.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
                }
            }

            XSSFCell cell5 = row.createCell(column+5);
            if (i==0){
                cell5.setCellStyle(titleModel.cellCommonStyle2(workbook));
//                cell5.setCellValue((ty[i]-ly[i])/ly[i]*100);
                if (ly[i]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((ty[i]-ly[i])/ly[i]*100);
                }
            }else if (i==dept.size()-1){
                cell5.setCellStyle(titleModel.cellCommonStyle3(workbook));
//                cell5.setCellValue((ty1-ly1)/ly1*100);
                if (ly1==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((ty1-ly1)/ly1*100);
                }
            }else {
                cell5.setCellStyle(titleModel.cellCommonStyle3(workbook));
                if (ly[i]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((ty[i]-ly[i])/ly[i]*100);
                }
            }
        }
    }

    public void outputMB1Excel(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                          double[] tytm,  double[] ty,  double[] lytm,  double[] ly,
                          double[] tytm1, double[] ty1, double[] lytm1, double[] ly1,
                          double[] tytm2, double[] ty2, double[] lytm2, double[] ly2,
                          double[] tytm3, double[] ty3, double[] lytm3, double[] ly3){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
            deptCell.setCellValue(dept.get(i));

            //确认收入
            //本月
            XSSFCell cell1 = row.createCell(2);
            cellStyle1(i,cell1,workbook,titleModel);
            cell1.setCellValue(tytm[i]);

            //本年
            XSSFCell cell2 = row.createCell(3);
            cellStyle1(i,cell2,workbook,titleModel);
            cell2.setCellValue(ty[i]);

            //本月同期
            XSSFCell cell3 = row.createCell(4);
            cellStyle1(i,cell3,workbook,titleModel);
            cell3.setCellValue(lytm[i]);

            //本年同期
            XSSFCell cell4 = row.createCell(5);
            cellStyle1(i,cell4,workbook,titleModel);
            cell4.setCellValue(ly[i]);

            //本月同比
            XSSFCell cell5 = row.createCell(6);
            cellStyle1(i,cell5,workbook,titleModel);
            if (lytm[i]==0){
                cell5.setCellValue(0);
            }else {
                cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
            }

            //本年同比
            XSSFCell cell6 = row.createCell(7);
            cellStyle1(i,cell6,workbook,titleModel);
            if (ly[i]==0){
                cell6.setCellValue(0);
            }else {
                cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
            }


            //挂账收入
            //本月
            XSSFCell cell7 = row.createCell(8);
            cellStyle1(i,cell7,workbook,titleModel);
            cell7.setCellValue(tytm1[i]);

            //本年
            XSSFCell cell8 = row.createCell(9);
            cellStyle1(i,cell8,workbook,titleModel);
            cell8.setCellValue(ty1[i]);

            //本月同期
            XSSFCell cell9 = row.createCell(10);
            cellStyle1(i,cell9,workbook,titleModel);
            cell9.setCellValue(lytm1[i]);

            //本年同期
            XSSFCell cell10 = row.createCell(11);
            cellStyle1(i,cell10,workbook,titleModel);
            cell10.setCellValue(ly1[i]);

            //本月同比
            XSSFCell cell11 = row.createCell(12);
            cellStyle1(i,cell11,workbook,titleModel);
            if (lytm1[i]==0){
                cell11.setCellValue(0);
            }else {
                cell11.setCellValue((tytm1[i]-lytm1[i])/lytm1[i]*100);
            }

            //本年同比
            XSSFCell cell12 = row.createCell(13);
            cellStyle1(i,cell12,workbook,titleModel);
            if (ly1[i]==0){
                cell12.setCellValue(0);
            }else {
                cell12.setCellValue((ty1[i]-ly1[i])/ly1[i]*100);
            }

            //主营业务成本
            XSSFCell cell13 = row.createCell(14);
            cellStyle1(i,cell13,workbook,titleModel);
            cell13.setCellValue(tytm2[i]);

            XSSFCell cell14 = row.createCell(15);
            cellStyle1(i,cell14,workbook,titleModel);
            cell14.setCellValue(ty2[i]);

            XSSFCell cell15 = row.createCell(16);
            cellStyle1(i,cell15,workbook,titleModel);
            cell15.setCellValue(lytm2[i]);

            XSSFCell cell16 = row.createCell(17);
            cellStyle1(i,cell16,workbook,titleModel);
            cell16.setCellValue(ly2[i]);

            XSSFCell cell17 = row.createCell(18);
            cellStyle1(i,cell17,workbook,titleModel);
            if (lytm2[i]==0){
                cell17.setCellValue(0);
            }else {
                cell17.setCellValue((tytm2[i]-lytm2[i])/lytm2[i]*100);
            }

            XSSFCell cell18 = row.createCell(19);
            cellStyle1(i,cell18,workbook,titleModel);
            if (ly2[i]==0){
                cell18.setCellValue(0);
            }else {
                cell18.setCellValue((ty2[i]-ly2[i])/ly2[i]*100);
            }

            XSSFCell cell19 = row.createCell(20);
            cellStyle1(i,cell19,workbook,titleModel);
            cell19.setCellValue(tytm3[i]);

            XSSFCell cell20 = row.createCell(21);
            cellStyle1(i,cell20,workbook,titleModel);
            cell20.setCellValue(ty3[i]);

            XSSFCell cell21 = row.createCell(22);
            cellStyle1(i,cell21,workbook,titleModel);
            cell21.setCellValue(lytm3[i]);

            XSSFCell cell22 = row.createCell(23);
            cellStyle1(i,cell22,workbook,titleModel);
            cell22.setCellValue(ly3[i]);
        }
    }

    public void outputMB1Excel(String dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                               double tytm,  double ty,  double lytm,  double ly,
                               double tytm1, double ty1, double lytm1, double ly1,
                               double tytm2, double ty2, double lytm2, double ly2,
                               double tytm3, double ty3, double lytm3, double ly3){

        TitleModel titleModel = new TitleModel();

        XSSFRow row = sheet.createRow(index);

            //部门列
            XSSFCell deptCell = row.createCell(1);
            deptCell.setCellStyle(titleModel.cellCommonStyle(workbook));
            deptCell.setCellValue(dept);

            //确认收入
            //本月
            XSSFCell cell1 = row.createCell(2);
            cell1.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell1.setCellValue(tytm);

            //本年
            XSSFCell cell2 = row.createCell(3);
            cell2.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell2.setCellValue(ty);

            //本月同期
            XSSFCell cell3 = row.createCell(4);
            cell3.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell3.setCellValue(lytm);

            //本年同期
            XSSFCell cell4 = row.createCell(5);
            cell4.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell4.setCellValue(ly);

            //本月同比
            XSSFCell cell5 = row.createCell(6);
            cell5.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (lytm==0){
                cell5.setCellValue(0);
            }else {
                cell5.setCellValue((tytm-lytm)/lytm*100);
            }

            //本年同比
            XSSFCell cell6 = row.createCell(7);
            cell6.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (ly==0){
                cell6.setCellValue(0);
            }else {
                cell6.setCellValue((ty-ly)/ly*100);
            }


            //挂账收入
            //本月
            XSSFCell cell7 = row.createCell(8);
            cell7.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell7.setCellValue(tytm1);

            //本年
            XSSFCell cell8 = row.createCell(9);
            cell8.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell8.setCellValue(ty1);

            //本月同期
            XSSFCell cell9 = row.createCell(10);
            cell9.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell9.setCellValue(lytm1);

            //本年同期
            XSSFCell cell10 = row.createCell(11);
            cell10.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell10.setCellValue(ly1);

            //本月同比
            XSSFCell cell11 = row.createCell(12);
            cell11.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (lytm1==0){
                cell11.setCellValue(0);
            }else {
                cell11.setCellValue((tytm1-lytm1)/lytm1*100);
            }

            //本年同比
            XSSFCell cell12 = row.createCell(13);
            cell12.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (ly1==0){
                cell12.setCellValue(0);
            }else {
                cell12.setCellValue((ty1-ly1)/ly1*100);
            }

            //主营业务成本
            XSSFCell cell13 = row.createCell(14);
            cell13.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell13.setCellValue(tytm2);

            XSSFCell cell14 = row.createCell(15);
            cell14.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell14.setCellValue(ty2);

            XSSFCell cell15 = row.createCell(16);
            cell15.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell15.setCellValue(lytm2);

            XSSFCell cell16 = row.createCell(17);
            cell16.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell16.setCellValue(ly2);

            XSSFCell cell17 = row.createCell(18);
            cell17.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (lytm2==0){
                cell17.setCellValue(0);
            }else {
                cell17.setCellValue((tytm2-lytm2)/lytm2*100);
            }

            XSSFCell cell18 = row.createCell(19);
            cell18.setCellStyle(titleModel.cellCommonStyle2(workbook));
            if (ly2==0){
                cell18.setCellValue(0);
            }else {
                cell18.setCellValue((ty2-ly2)/ly2*100);
            }

            XSSFCell cell19 = row.createCell(20);
            cell19.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell19.setCellValue(tytm3);

            XSSFCell cell20 = row.createCell(21);
            cell20.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell20.setCellValue(ty3);

            XSSFCell cell21 = row.createCell(22);
            cell21.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell21.setCellValue(lytm3);

            XSSFCell cell22 = row.createCell(23);
            cell22.setCellStyle(titleModel.cellCommonStyle2(workbook));
            cell22.setCellValue(ly3);

    }

    public void outputMB1Excel(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                               double[] tytm,  double[] ty,  double[] lytm,  double[] ly,
                               double[] tytm1, double[] ty1, double[] lytm1, double[] ly1,
                               double[] tytm2, double[] ty2, double[] lytm2, double[] ly2,
                               double[] tytm3, double[] ty3, double[] lytm3, double[] ly3,
                               double tytm4, double ty4, double lytm4, double ly4){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);
            if (i==1){
                //本月
                XSSFCell cell1 = row.createCell(14);
                cellStyle1(i,cell1,workbook,titleModel);
                cell1.setCellValue(tytm4);

                //本年
                XSSFCell cell2 = row.createCell(15);
                cellStyle1(i,cell2,workbook,titleModel);
                cell2.setCellValue(ty4);

                //本月同期
                XSSFCell cell3 = row.createCell(16);
                cellStyle1(i,cell3,workbook,titleModel);
                cell3.setCellValue(lytm4);

                //本年同期
                XSSFCell cell4 = row.createCell(17);
                cellStyle1(i,cell4,workbook,titleModel);
                cell4.setCellValue(ly4);

                //本月同比
                XSSFCell cell5 = row.createCell(18);
                cellStyle1(i,cell5,workbook,titleModel);
                if (lytm4==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm4-lytm4)/lytm4*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(19);
                cellStyle1(i,cell6,workbook,titleModel);
                if (ly4==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty4-ly4)/ly4*100);
                }
            }else if (i==0){
                //确认收入
                //本月
                XSSFCell cell1 = row.createCell(2);
                cellStyle1(i,cell1,workbook,titleModel);
                cell1.setCellValue(tytm[i]);

                //本年
                XSSFCell cell2 = row.createCell(3);
                cellStyle1(i,cell2,workbook,titleModel);
                cell2.setCellValue(ty[i]);

                //本月同期
                XSSFCell cell3 = row.createCell(4);
                cellStyle1(i,cell3,workbook,titleModel);
                cell3.setCellValue(lytm[i]);

                //本年同期
                XSSFCell cell4 = row.createCell(5);
                cellStyle1(i,cell4,workbook,titleModel);
                cell4.setCellValue(ly[i]);

                //本月同比
                XSSFCell cell5 = row.createCell(6);
                cellStyle1(i,cell5,workbook,titleModel);
                if (lytm[i]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(7);
                cellStyle1(i,cell6,workbook,titleModel);
                if (ly[i]==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
                }


                //挂账收入
                //本月
                XSSFCell cell7 = row.createCell(8);
                cellStyle1(i,cell7,workbook,titleModel);
                cell7.setCellValue(tytm1[i]);

                //本年
                XSSFCell cell8 = row.createCell(9);
                cellStyle1(i,cell8,workbook,titleModel);
                cell8.setCellValue(ty1[i]);

                //本月同期
                XSSFCell cell9 = row.createCell(10);
                cellStyle1(i,cell9,workbook,titleModel);
                cell9.setCellValue(lytm1[i]);

                //本年同期
                XSSFCell cell10 = row.createCell(11);
                cellStyle1(i,cell10,workbook,titleModel);
                cell10.setCellValue(ly1[i]);

                //本月同比
                XSSFCell cell11 = row.createCell(12);
                cellStyle1(i,cell11,workbook,titleModel);
                if (lytm1[i]==0){
                    cell11.setCellValue(0);
                }else {
                    cell11.setCellValue((tytm1[i]-lytm1[i])/lytm1[i]*100);
                }

                //本年同比
                XSSFCell cell12 = row.createCell(13);
                cellStyle1(i,cell12,workbook,titleModel);
                if (ly1[i]==0){
                    cell12.setCellValue(0);
                }else {
                    cell12.setCellValue((ty1[i]-ly1[i])/ly1[i]*100);
                }

                //主营业务成本
                XSSFCell cell13 = row.createCell(14);
                cellStyle1(i,cell13,workbook,titleModel);
                cell13.setCellValue(tytm2[i]);

                XSSFCell cell14 = row.createCell(15);
                cellStyle1(i,cell14,workbook,titleModel);
                cell14.setCellValue(ty2[i]);

                XSSFCell cell15 = row.createCell(16);
                cellStyle1(i,cell15,workbook,titleModel);
                cell15.setCellValue(lytm2[i]);

                XSSFCell cell16 = row.createCell(17);
                cellStyle1(i,cell16,workbook,titleModel);
                cell16.setCellValue(ly2[i]);

                XSSFCell cell17 = row.createCell(18);
                cellStyle1(i,cell17,workbook,titleModel);
                if (lytm2[i]==0){
                    cell17.setCellValue(0);
                }else {
                    cell17.setCellValue((tytm2[i]-lytm2[i])/lytm2[i]*100);
                }

                XSSFCell cell18 = row.createCell(19);
                cellStyle1(i,cell18,workbook,titleModel);
                if (ly2[i]==0){
                    cell18.setCellValue(0);
                }else {
                    cell18.setCellValue((ty2[i]-ly2[i])/ly2[i]*100);
                }

                XSSFCell cell19 = row.createCell(20);
                cellStyle1(i,cell19,workbook,titleModel);
                cell19.setCellValue(tytm3[i]);

                XSSFCell cell20 = row.createCell(21);
                cellStyle1(i,cell20,workbook,titleModel);
                cell20.setCellValue(ty3[i]);

                XSSFCell cell21 = row.createCell(22);
                cellStyle1(i,cell21,workbook,titleModel);
                cell21.setCellValue(lytm3[i]);

                XSSFCell cell22 = row.createCell(23);
                cellStyle1(i,cell22,workbook,titleModel);
                cell22.setCellValue(ly3[i]);
            }else {
                //确认收入
                //本月
                XSSFCell cell1 = row.createCell(2);
                cellStyle1(i-1,cell1,workbook,titleModel);
                cell1.setCellValue(tytm[i-1]);

                //本年
                XSSFCell cell2 = row.createCell(3);
                cellStyle1(i-1,cell2,workbook,titleModel);
                cell2.setCellValue(ty[i-1]);

                //本月同期
                XSSFCell cell3 = row.createCell(4);
                cellStyle1(i-1,cell3,workbook,titleModel);
                cell3.setCellValue(lytm[i-1]);

                //本年同期
                XSSFCell cell4 = row.createCell(5);
                cellStyle1(i-1,cell4,workbook,titleModel);
                cell4.setCellValue(ly[i-1]);

                //本月同比
                XSSFCell cell5 = row.createCell(6);
                cellStyle1(i-1,cell5,workbook,titleModel);
                if (lytm[i-1]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm[i-1]-lytm[i-1])/lytm[i-1]*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(7);
                cellStyle1(i-1,cell6,workbook,titleModel);
                if (ly[i-1]==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty[i-1]-ly[i-1])/ly[i-1]*100);
                }


                //挂账收入
                //本月
                XSSFCell cell7 = row.createCell(8);
                cellStyle1(i-1,cell7,workbook,titleModel);
                cell7.setCellValue(tytm1[i-1]);

                //本年
                XSSFCell cell8 = row.createCell(9);
                cellStyle1(i-1,cell8,workbook,titleModel);
                cell8.setCellValue(ty1[i-1]);

                //本月同期
                XSSFCell cell9 = row.createCell(10);
                cellStyle1(i-1,cell9,workbook,titleModel);
                cell9.setCellValue(lytm1[i-1]);

                //本年同期
                XSSFCell cell10 = row.createCell(11);
                cellStyle1(i-1,cell10,workbook,titleModel);
                cell10.setCellValue(ly1[i-1]);

                //本月同比
                XSSFCell cell11 = row.createCell(12);
                cellStyle1(i-1,cell11,workbook,titleModel);
                if (lytm1[i-1]==0){
                    cell11.setCellValue(0);
                }else {
                    cell11.setCellValue((tytm1[i-1]-lytm1[i-1])/lytm1[i-1]*100);
                }

                //本年同比
                XSSFCell cell12 = row.createCell(13);
                cellStyle1(i-1,cell12,workbook,titleModel);
                if (ly1[i-1]==0){
                    cell12.setCellValue(0);
                }else {
                    cell12.setCellValue((ty1[i-1]-ly1[i-1])/ly1[i-1]*100);
                }

                //主营业务成本
                XSSFCell cell13 = row.createCell(14);
                cellStyle1(i-1,cell13,workbook,titleModel);
                cell13.setCellValue(tytm2[i-1]);

                XSSFCell cell14 = row.createCell(15);
                cellStyle1(i-1,cell14,workbook,titleModel);
                cell14.setCellValue(ty2[i-1]);

                XSSFCell cell15 = row.createCell(16);
                cellStyle1(i-1,cell15,workbook,titleModel);
                cell15.setCellValue(lytm2[i-1]);

                XSSFCell cell16 = row.createCell(17);
                cellStyle1(i-1,cell16,workbook,titleModel);
                cell16.setCellValue(ly2[i-1]);

                XSSFCell cell17 = row.createCell(18);
                cellStyle1(i-1,cell17,workbook,titleModel);
                if (lytm2[i-1]==0){
                    cell17.setCellValue(0);
                }else {
                    cell17.setCellValue((tytm2[i-1]-lytm2[i-1])/lytm2[i-1]*100);
                }

                XSSFCell cell18 = row.createCell(19);
                cellStyle1(i,cell18,workbook,titleModel);
                if (ly2[i-1]==0){
                    cell18.setCellValue(0);
                }else {
                    cell18.setCellValue((ty2[i-1]-ly2[i-1])/ly2[i-1]*100);
                }

                XSSFCell cell19 = row.createCell(20);
                cellStyle1(i,cell19,workbook,titleModel);
                cell19.setCellValue(tytm3[i-1]);

                XSSFCell cell20 = row.createCell(21);
                cellStyle1(i,cell20,workbook,titleModel);
                cell20.setCellValue(ty3[i-1]);

                XSSFCell cell21 = row.createCell(22);
                cellStyle1(i,cell21,workbook,titleModel);
                cell21.setCellValue(lytm3[i-1]);

                XSSFCell cell22 = row.createCell(23);
                cellStyle1(i,cell22,workbook,titleModel);
                cell22.setCellValue(ly3[i-1]);
            }

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
            deptCell.setCellValue(dept.get(i));


            
        }
    }

    public void outputCFSExcel(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                               double[] tytm,  double[] ty,  double[] lytm,  double[] ly,
                               double[] tytm1, double[] ty1, double[] lytm1, double[] ly1,
                               double[] tytm2, double[] ty2, double[] lytm2, double[] ly2){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
            deptCell.setCellValue(dept.get(i));

            //确认收入
            //本月
            XSSFCell cell1 = row.createCell(2);
            cellStyle1(i,cell1,workbook,titleModel);
            cell1.setCellValue(tytm[i]);

            //本年
            XSSFCell cell2 = row.createCell(3);
            cellStyle1(i,cell2,workbook,titleModel);
            cell2.setCellValue(ty[i]);

            //本月同期
            XSSFCell cell3 = row.createCell(4);
            cellStyle1(i,cell3,workbook,titleModel);
            cell3.setCellValue(lytm[i]);

            //本年同期
            XSSFCell cell4 = row.createCell(5);
            cellStyle1(i,cell4,workbook,titleModel);
            cell4.setCellValue(ly[i]);

            //本月同比
            XSSFCell cell5 = row.createCell(6);
            cellStyle1(i,cell5,workbook,titleModel);
            if (lytm[i]==0){
                cell5.setCellValue(0);
            }else {
                cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
            }

            //本年同比
            XSSFCell cell6 = row.createCell(7);
            cellStyle1(i,cell6,workbook,titleModel);
            if (ly[i]==0){
                cell6.setCellValue(0);
            }else {
                cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
            }


            //挂账收入
            //本月
            XSSFCell cell7 = row.createCell(8);
            cellStyle1(i,cell7,workbook,titleModel);
            cell7.setCellValue(tytm1[i]);

            //本年
            XSSFCell cell8 = row.createCell(9);
            cellStyle1(i,cell8,workbook,titleModel);
            cell8.setCellValue(ty1[i]);

            //本月同期
            XSSFCell cell9 = row.createCell(10);
            cellStyle1(i,cell9,workbook,titleModel);
            cell9.setCellValue(lytm1[i]);

            //本年同期
            XSSFCell cell10 = row.createCell(11);
            cellStyle1(i,cell10,workbook,titleModel);
            cell10.setCellValue(ly1[i]);

            //本月同比
            XSSFCell cell11 = row.createCell(12);
            cellStyle1(i,cell11,workbook,titleModel);
            if (lytm1[i]==0){
                cell11.setCellValue(0);
            }else {
                cell11.setCellValue((tytm1[i]-lytm1[i])/lytm1[i]*100);
            }

            //本年同比
            XSSFCell cell12 = row.createCell(13);
            cellStyle1(i,cell12,workbook,titleModel);
            if (ly1[i]==0){
                cell12.setCellValue(0);
            }else {
                cell12.setCellValue((ty1[i]-ly1[i])/ly1[i]*100);
            }

            //主营业务成本
            XSSFCell cell13 = row.createCell(14);
            cellStyle1(i,cell13,workbook,titleModel);
            cell13.setCellValue(tytm2[i]);

            XSSFCell cell14 = row.createCell(15);
            cellStyle1(i,cell14,workbook,titleModel);
            cell14.setCellValue(ty2[i]);

            XSSFCell cell15 = row.createCell(16);
            cellStyle1(i,cell15,workbook,titleModel);
            cell15.setCellValue(lytm2[i]);

            XSSFCell cell16 = row.createCell(17);
            cellStyle1(i,cell16,workbook,titleModel);
            cell16.setCellValue(ly2[i]);

            XSSFCell cell17 = row.createCell(18);
            cellStyle1(i,cell17,workbook,titleModel);
            if (lytm2[i]==0){
                cell17.setCellValue(0);
            }else {
                cell17.setCellValue((tytm2[i]-lytm2[i])/lytm2[i]*100);
            }

            XSSFCell cell18 = row.createCell(19);
            cellStyle1(i,cell18,workbook,titleModel);
            if (ly2[i]==0){
                cell18.setCellValue(0);
            }else {
                cell18.setCellValue((ty2[i]-ly2[i])/ly2[i]*100);
            }
        }
    }

    public void outputSalaryExcel(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                               double[] tytm,  double[] ty,  double[] lytm,  double[] ly,
                               double tytm2, double ty2, double lytm2, double ly2){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);
            if (i==1){
                //本月
                XSSFCell cell1 = row.createCell(2);
                cellStyle1(i,cell1,workbook,titleModel);
                cell1.setCellValue(tytm2);

                //本年
                XSSFCell cell2 = row.createCell(3);
                cellStyle1(i,cell2,workbook,titleModel);
                cell2.setCellValue(ty2);

                //本月同期
                XSSFCell cell3 = row.createCell(4);
                cellStyle1(i,cell3,workbook,titleModel);
                cell3.setCellValue(lytm2);

                //本年同期
                XSSFCell cell4 = row.createCell(5);
                cellStyle1(i,cell4,workbook,titleModel);
                cell4.setCellValue(ly2);

                //本月同比
                XSSFCell cell5 = row.createCell(6);
                cellStyle1(i,cell5,workbook,titleModel);
                if (lytm2==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm2-lytm2)/lytm2*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(7);
                cellStyle1(i,cell6,workbook,titleModel);
                if (ly2==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty2-ly2)/ly2*100);
                }
            }else if (i==0){
                //确认收入
                //本月
                XSSFCell cell1 = row.createCell(2);
                cellStyle1(i,cell1,workbook,titleModel);
                cell1.setCellValue(tytm[i]);

                //本年
                XSSFCell cell2 = row.createCell(3);
                cellStyle1(i,cell2,workbook,titleModel);
                cell2.setCellValue(ty[i]);

                //本月同期
                XSSFCell cell3 = row.createCell(4);
                cellStyle1(i,cell3,workbook,titleModel);
                cell3.setCellValue(lytm[i]);

                //本年同期
                XSSFCell cell4 = row.createCell(5);
                cellStyle1(i,cell4,workbook,titleModel);
                cell4.setCellValue(ly[i]);

                //本月同比
                XSSFCell cell5 = row.createCell(6);
                cellStyle1(i,cell5,workbook,titleModel);
                if (lytm[i]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(7);
                cellStyle1(i,cell6,workbook,titleModel);
                if (ly[i]==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
                }
            }else {
                //确认收入
                //本月
                XSSFCell cell1 = row.createCell(2);
                cellStyle1(i-1,cell1,workbook,titleModel);
                cell1.setCellValue(tytm[i-1]);

                //本年
                XSSFCell cell2 = row.createCell(3);
                cellStyle1(i-1,cell2,workbook,titleModel);
                cell2.setCellValue(ty[i-1]);

                //本月同期
                XSSFCell cell3 = row.createCell(4);
                cellStyle1(i-1,cell3,workbook,titleModel);
                cell3.setCellValue(lytm[i-1]);

                //本年同期
                XSSFCell cell4 = row.createCell(5);
                cellStyle1(i-1,cell4,workbook,titleModel);
                cell4.setCellValue(ly[i-1]);

                //本月同比
                XSSFCell cell5 = row.createCell(6);
                cellStyle1(i-1,cell5,workbook,titleModel);
                if (lytm[i-1]==0){
                    cell5.setCellValue(0);
                }else {
                    cell5.setCellValue((tytm[i-1]-lytm[i-1])/lytm[i-1]*100);
                }

                //本年同比
                XSSFCell cell6 = row.createCell(7);
                cellStyle1(i-1,cell6,workbook,titleModel);
                if (ly[i-1]==0){
                    cell6.setCellValue(0);
                }else {
                    cell6.setCellValue((ty[i-1]-ly[i-1])/ly[i-1]*100);
                }
            }

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
            deptCell.setCellValue(dept.get(i));



        }
    }



    public void outputMBI(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                                         double[] tytm,  double[] ty,  double[] lytm,  double[] ly,
                                         double[] tytm1, double[] ty1, double[] lytm1, double[] ly1,
                                         double[] tytm2, double[] ty2, double[] lytm2, double[] ly2){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);

            XSSFCell cell = row.createCell(1);
            cellStyle(i,cell,workbook,titleModel);
            cell.setCellValue(dept.get(i));

            //确认收入
            XSSFCell cell1 = row.createCell(2);
            cellStyle1(i,cell1,workbook,titleModel);
            cell1.setCellValue(tytm[i]);

            XSSFCell cell2 = row.createCell(3);
            cellStyle1(i,cell2,workbook,titleModel);
            cell2.setCellValue(ty[i]);

            XSSFCell cell3 = row.createCell(4);
            cellStyle1(i,cell3,workbook,titleModel);
            cell3.setCellValue(lytm[i]);

            XSSFCell cell4 = row.createCell(5);
            cellStyle1(i,cell4,workbook,titleModel);
            cell4.setCellValue(ly[i]);

            XSSFCell cell5 = row.createCell(6);
            cellStyle1(i,cell5,workbook,titleModel);
            if (lytm[i]==0){
                cell5.setCellValue(0);
            }else {
                cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
            }

            XSSFCell cell6 = row.createCell(7);
            cellStyle1(i,cell6,workbook,titleModel);
            if (ly[i]==0){
                cell6.setCellValue(0);
            }else {
                cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
            }

            //挂账收入
            XSSFCell cell7 = row.createCell(8);
            cellStyle1(i,cell7,workbook,titleModel);
            cell7.setCellValue(tytm1[i]);

            XSSFCell cell8 = row.createCell(9);
            cellStyle1(i,cell8,workbook,titleModel);
            cell8.setCellValue(ty1[i]);

            XSSFCell cell9 = row.createCell(10);
            cellStyle1(i,cell9,workbook,titleModel);
            cell9.setCellValue(lytm1[i]);

            XSSFCell cell10 = row.createCell(11);
            cellStyle1(i,cell10,workbook,titleModel);
            cell10.setCellValue(ly1[i]);

            XSSFCell cell11 = row.createCell(12);
            cellStyle1(i,cell11,workbook,titleModel);
            if (lytm1[i]==0){
                cell11.setCellValue(0);
            }else {
                cell11.setCellValue((tytm1[i]-lytm1[i])/lytm1[i]*100);
            }

            XSSFCell cell12 = row.createCell(13);
            cellStyle1(i,cell12,workbook,titleModel);
            if (ly1[i]==0){
                cell12.setCellValue(0);
            }else {
                cell12.setCellValue((ty1[i]-ly1[i])/ly1[i]*100);
            }

            //毛利率
            XSSFCell cell13 = row.createCell(14);
            cellStyle1(i,cell13,workbook,titleModel);
            cell13.setCellValue(tytm2[i]);

            XSSFCell cell14 = row.createCell(15);
            cellStyle1(i,cell14,workbook,titleModel);
            cell14.setCellValue(ty2[i]);

            XSSFCell cell15 = row.createCell(16);
            cellStyle1(i,cell15,workbook,titleModel);
            cell15.setCellValue(lytm2[i]);

            XSSFCell cell16 = row.createCell(17);
            cellStyle1(i,cell16,workbook,titleModel);
            cell16.setCellValue(ly2[i]);

            XSSFCell cell17 = row.createCell(18);
            cellStyle1(i,cell17,workbook,titleModel);
            if (lytm2[i]==0){
                cell17.setCellValue(0);
            }else {
                cell17.setCellValue((tytm2[i]-lytm2[i])/lytm2[i]*100);
            }

            XSSFCell cell18 = row.createCell(19);
            cellStyle1(i,cell18,workbook,titleModel);
            if (ly2[i]==0){
                cell18.setCellValue(0);
            }else {
                cell18.setCellValue((ty2[i]-ly2[i])/ly2[i]*100);
            }
        }
    }

    public void outputMBI(String dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                          double tytm,  double ty,  double lytm,  double ly,
                          double tytm1, double ty1, double lytm1, double ly1,
                          double tytm2, double ty2, double lytm2, double ly2){

        TitleModel titleModel = new TitleModel();

        XSSFRow row = sheet.createRow(index);


        XSSFCell deptCell = row.createCell(1);
        deptCell.setCellStyle(titleModel.cellCommonStyle(workbook));
        deptCell.setCellValue(dept);


        XSSFCell cell1 = row.createCell(2);
        cell1.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell1.setCellValue(tytm);

        XSSFCell cell2 = row.createCell(3);
        cell2.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell2.setCellValue(ty);

        XSSFCell cell3 = row.createCell(4);
        cell3.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell3.setCellValue(lytm);

        XSSFCell cell4 = row.createCell(5);
        cell4.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell4.setCellValue(ly);

        XSSFCell cell5 = row.createCell(6);
        cell5.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (lytm==0){
            cell5.setCellValue(0);
        }else {
            cell5.setCellValue((tytm-lytm)/lytm*100);
        }

        XSSFCell cell6 = row.createCell(7);
        cell6.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (ly==0){
            cell6.setCellValue(0);
        }else {
            cell6.setCellValue((ty-ly)/ly*100);
        }


        XSSFCell cell7 = row.createCell(8);
        cell7.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell7.setCellValue(tytm1);

        XSSFCell cell8 = row.createCell(9);
        cell8.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell8.setCellValue(ty1);

        XSSFCell cell9 = row.createCell(10);
        cell9.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell9.setCellValue(lytm1);

        XSSFCell cell10 = row.createCell(11);
        cell10.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell10.setCellValue(ly1);

        XSSFCell cell11 = row.createCell(12);
        cell11.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (lytm1==0){
            cell11.setCellValue(0);
        }else {
            cell11.setCellValue((tytm1-lytm1)/lytm1*100);
        }

        XSSFCell cell12 = row.createCell(13);
        cell12.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (ly1==0){
            cell12.setCellValue(0);
        }else {
            cell12.setCellValue((ty1-ly1)/ly1*100);
        }


        XSSFCell cell13 = row.createCell(14);
        cell13.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell13.setCellValue(tytm2);

        XSSFCell cell14 = row.createCell(15);
        cell14.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell14.setCellValue(ty2);

        XSSFCell cell15 = row.createCell(16);
        cell15.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell15.setCellValue(lytm2);

        XSSFCell cell16 = row.createCell(17);
        cell16.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell16.setCellValue(ly2);

        XSSFCell cell17 = row.createCell(18);
        cell17.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (lytm2==0){
            cell17.setCellValue(0);
        }else {
            cell17.setCellValue((tytm2-lytm2)/lytm2*100);
        }

        XSSFCell cell18 = row.createCell(19);
        cell18.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (ly2==0){
            cell18.setCellValue(0);
        }else {
            cell18.setCellValue((ty2-ly2)/ly2*100);
        }

    }

    public void outputSome(List<String> dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                           double[] tytm, double[] ty, double[] lytm, double[] ly){

        TitleModel titleModel = new TitleModel();

        for (int i = 0; i < dept.size(); i++) {
            XSSFRow row = sheet.createRow(i+index);

            //部门列
            XSSFCell deptCell = row.createCell(1);
            cellStyle(i,deptCell,workbook,titleModel);
            deptCell.setCellValue(dept.get(i));

            //确认收入
            //本月
            XSSFCell cell1 = row.createCell(2);
            cellStyle1(i,cell1,workbook,titleModel);
            cell1.setCellValue(tytm[i]);

            //本年
            XSSFCell cell2 = row.createCell(3);
            cellStyle1(i,cell2,workbook,titleModel);
            cell2.setCellValue(ty[i]);

            //本月同期
            XSSFCell cell3 = row.createCell(4);
            cellStyle1(i,cell3,workbook,titleModel);
            cell3.setCellValue(lytm[i]);

            //本年同期
            XSSFCell cell4 = row.createCell(5);
            cellStyle1(i,cell4,workbook,titleModel);
            cell4.setCellValue(ly[i]);

            //本月同比
            XSSFCell cell5 = row.createCell(6);
            cellStyle1(i,cell5,workbook,titleModel);
            if (lytm[i]==0){
                cell5.setCellValue(0);
            }else {
                cell5.setCellValue((tytm[i]-lytm[i])/lytm[i]*100);
            }

            //本年同比
            XSSFCell cell6 = row.createCell(7);
            cellStyle1(i,cell6,workbook,titleModel);
            if (ly[i]==0){
                cell6.setCellValue(0);
            }else {
                cell6.setCellValue((ty[i]-ly[i])/ly[i]*100);
            }
        }
    }

    public void outputSingle(String dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,int column,
                             double tytm, double ty, double lytm, double ly){
        TitleModel titleModel = new TitleModel();

        XSSFRow row = sheet.createRow(index);

        XSSFCell cell = row.createCell(1);
        cell.setCellStyle(titleModel.cellCommonStyle(workbook));
        cell.setCellValue(dept);

        XSSFCell cell1 = row.createCell(column);
        cell1.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell1.setCellValue(tytm);

        XSSFCell cell2 = row.createCell(column+1);
        cell2.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell2.setCellValue(ty);

        XSSFCell cell3 = row.createCell(column+2);
        cell3.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell3.setCellValue(lytm);

        XSSFCell cell4 = row.createCell(column+3);
        cell4.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell4.setCellValue(ly);

        XSSFCell cell5 = row.createCell(column+4);
        cell5.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (lytm==0){
            cell5.setCellValue(0);
        }else {
            cell5.setCellValue((tytm-lytm)/lytm*100);
        }

        XSSFCell cell6 = row.createCell(column+5);
        cell6.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (ly==0){
            cell6.setCellValue(0);
        }else {
            cell6.setCellValue((ty-ly)/ly*100);
        }
    }

    public void outputSome(String dept, XSSFWorkbook workbook, XSSFSheet sheet, int index,
                           double tytm, double ty, double lytm, double ly){
        TitleModel titleModel = new TitleModel();

        XSSFRow row = sheet.createRow(index);

        XSSFCell cell = row.createCell(1);
        cell.setCellStyle(titleModel.cellCommonStyle(workbook));
        cell.setCellValue(dept);

        XSSFCell cell1 = row.createCell(8);
        cell1.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell1.setCellValue(tytm);

        XSSFCell cell2 = row.createCell(9);
        cell2.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell2.setCellValue(ty);

        XSSFCell cell3 = row.createCell(10);
        cell3.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell3.setCellValue(lytm);

        XSSFCell cell4 = row.createCell(11);
        cell4.setCellStyle(titleModel.cellCommonStyle2(workbook));
        cell4.setCellValue(ly);

        XSSFCell cell5 = row.createCell(12);
        cell5.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (lytm==0){
            cell5.setCellValue(0);
        }else {
            cell5.setCellValue((tytm-lytm)/lytm*100);
        }

        XSSFCell cell6 = row.createCell(13);
        cell6.setCellStyle(titleModel.cellCommonStyle2(workbook));
        if (ly==0){
            cell6.setCellValue(0);
        }else {
            cell6.setCellValue((ty-ly)/ly*100);
        }
    }


    /**
     * 部门列单元格样式
     * @param i
     * @param cell
     * @param workbook
     * @param titleModel
     */
    public void cellStyle(int i,XSSFCell cell, XSSFWorkbook workbook,TitleModel titleModel){
        if (i==0){
            cell.setCellStyle(titleModel.cellCommonStyle(workbook));
        }else{
            cell.setCellStyle(titleModel.cellCommonStyle1(workbook));
        }
    }

    /**
     * 非部门列单元格样式
     * @param i
     * @param cell
     * @param workbook
     * @param titleModel
     */
    public void cellStyle1(int i,XSSFCell cell, XSSFWorkbook workbook,TitleModel titleModel){
        if (i==0){
            cell.setCellStyle(titleModel.cellCommonStyle2(workbook));
        }else{
            cell.setCellStyle(titleModel.cellCommonStyle3(workbook));
        }
    }
}
