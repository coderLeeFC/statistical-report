import cn.sljl.service.CFSReport;
import cn.sljl.service.MBIReport;
import cn.sljl.service.MainReport;
import cn.sljl.service.SalaryReport;
import cn.sljl.util.JdbcUtilsDbcp;
import cn.sljl.util.TitleModel;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * @author wangeqiu
 * @version 1.0
 * @date 2024/3/16 14:38
 */
public class Test {
    public static void main(String[] args) throws SQLException, IOException {
        Connection connection = JdbcUtilsDbcp.getConnection();
        XSSFWorkbook workbook = new XSSFWorkbook();
        TitleModel titleModel = new TitleModel();

        XSSFSheet sheetMain = workbook.createSheet("主表");//创建sheet
        new MainReport().mainMethod(connection,workbook, sheetMain,titleModel);//设置数据

//        XSSFSheet sheetMBI = workbook.createSheet("毛利率");//创建sheet
//        new MBIReport().mainMethod(connection,workbook, sheetMBI,titleModel);//设置数据
//
//        XSSFSheet sheetCFS = workbook.createSheet("现金流");//创建sheet
//        new CFSReport().mainMethod(connection,workbook, sheetCFS,titleModel);//设置数据
//
//        XSSFSheet sheetSR = workbook.createSheet("人工成本");//创建sheet
//        new SalaryReport().mainMethod(connection,workbook, sheetSR,titleModel);//设置数据

        FileOutputStream fileOutputStream = new FileOutputStream("D:\\综合统计报表x.xlsx");
        workbook.write(fileOutputStream);

        //释放资源
        fileOutputStream.flush();
        fileOutputStream.close();
        JdbcUtilsDbcp.release(connection,null,null);
    }
}
