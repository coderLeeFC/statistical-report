import cn.sljl.service.InvoiceReport;
import cn.sljl.service.MainReport;
import cn.sljl.util.JdbcUtilsDbcp;
import cn.sljl.util.TitleModel;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Connection;
import java.sql.SQLException;

/**
 * @author wangeqiu
 * @version 1.0
 * @date 2024/5/6 11:10
 */
public class InvoiceReportTest {
    public static void main(String[] args) throws SQLException, IOException {
        Connection connection = JdbcUtilsDbcp.getConnection();
        XSSFWorkbook workbook = new XSSFWorkbook();
        TitleModel titleModel = new TitleModel();

        XSSFSheet sheetMain = workbook.createSheet("生产经营部");//创建sheet
        new InvoiceReport().mainMethod(connection,workbook, sheetMain,titleModel);//设置数据

        FileOutputStream fileOutputStream = new FileOutputStream("D:\\生产经营部综合统计报表.xlsx");
        workbook.write(fileOutputStream);

        //释放资源
        fileOutputStream.flush();
        fileOutputStream.close();
        JdbcUtilsDbcp.release(connection,null,null);
    }
}
