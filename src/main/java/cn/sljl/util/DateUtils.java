package cn.sljl.util;

import java.util.Calendar;

/**
 * 对输入日期格式化
 *
 * @author wangeqiu
 * @version 1.0.0
 * @date 2024/03/20 15:35:45
 */
public class DateUtils {
    /**
     * 日历
     */
    Calendar calendar = Calendar.getInstance();

    /**
     * 获取去年上个月月末日期
     * @param endOfMonth
     * @return
     */
    public String getLastMonthEnd1(String endOfMonth){return (calendar.get(Calendar.YEAR) - 1)+getLastMonthEnd(endOfMonth).substring(4);}

    /**
     * 获取本年上个月月末日期
     * @param endOfMonth
     * @return
     */
    public String getLastMonthEnd(String endOfMonth){
        int i = Integer.parseInt(endOfMonth.substring(5, 7));
        return i<11?(endOfMonth.substring(0, 5)+"0"+(i-1)+"-31"):(endOfMonth.substring(0, 5)+(i-1)+"-31");
//        if (i<11){
//            return endOfMonth.substring(0, 5)+"0"+(i-1)+"-31";
//        }
//        return endOfMonth.substring(0, 5)+(i-1)+"-31";
    }

    /**
     * 获取本年期初日期
     * @param endOfMonth
     * @return
     */
    public String getStartOfYear(String endOfMonth){
        return endOfMonth.substring(0, 5) + "00-00";
    }

    /**
     * 获取去年期初日期
     * @param endOfMonth
     * @return
     */
    public String getStartOfLastYear(String endOfMonth){return (calendar.get(Calendar.YEAR) - 1)+getStartOfYear(endOfMonth).substring(4);}

    /**
     * 获取本月初日期
     *
     * @param endOfMonth 月底
     * @return {@link String }
     * @author wangeqiu
     * @date 2024/03/20 15:35:45
     */
    public String getBeginningOfMonth(String endOfMonth) {
        return endOfMonth.substring(0, 8) + "01";
    }

    /**
     * 获取本年初日期
     *
     * @param endOfMonth 月底
     * @return {@link String }
     * @author wangeqiu
     * @date 2024/03/20 15:35:45
     */
    public String getBeginningOfYear(String endOfMonth) {
        return endOfMonth.substring(0, 5) + "01-01";
    }

    /**
     * 获取去年月初
     *
     * @param endOfMonth 月底
     * @return {@link String }
     * @author wangeqiu
     * @date 2024/03/20 15:35:45
     */
    public String getBeginningOfMonth1(String endOfMonth) { return (calendar.get(Calendar.YEAR) - 1) + getBeginningOfMonth(endOfMonth).substring(4);}

    /**
     * 获取去年月末
     *
     * @param endOfMonth 月底
     * @return {@link String }
     * @author wangeqiu
     * @date 2024/03/20 15:35:46
     */
    public String getEndOfMonth(String endOfMonth) {return (calendar.get(Calendar.YEAR) - 1) + endOfMonth.substring(4,8)+"31";}

    /**
     * 获取去年年初
     *
     * @param endOfMonth 月底
     * @return {@link String }
     * @author wangeqiu
     * @date 2024/03/20 15:35:46
     */
    public String getBeginningOfYear1(String endOfMonth) { return (calendar.get(Calendar.YEAR) - 1) + getBeginningOfYear(endOfMonth).substring(4);}
}
