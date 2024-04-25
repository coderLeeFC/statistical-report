package cn.sljl.util;

import java.util.Arrays;
import java.util.List;

/**
 * 常量
 * @author wangeqiu
 * @version 1.0
 * @date 2024/3/16 9:05
 */
public class Constant {
    /**
     * 监理部门
     */
    public final static List<String> SLJL_BRANCH_OFFICE_NAME = Arrays.asList("工程监理","滨盘分公司","滨海分公司","东辛分公司","海洋分公司","电力分公司","山东管网","市直监理部","河北分公司","青岛分公司");
    public final static List<String> SLJL_DEPT_NUM = Arrays.asList(
            "0201", //滨盘
            "0202", //滨海
            "0203", //东辛
            "0204", //海洋
            "0205", //电力
            "0208", //山东管网
            "0209", //市直
            "0212");//河北
    public final static String ZBDL="0210";
    public final static String ZBDL_NAME="招标代理";

    public final static String ZJZX="0207";
    public final static String ZJZX_NAME="造价咨询";

    public final static String PGZX_DEPT="0215";

    //财务=财务+业务，人力=人力+人才储备，公司办=公司办+领导
    public final static List<String> SLJL_MANAGE_DEPT_NAME=Arrays.asList("机关部室","公司办公室","生产经营部","财务资产部（含中介机构费用）","人力资源部","总务管理部","市场开发部");
    public final static List<String> SLJL_MANAGE_DEPT=Arrays.asList(
            "0101", //领导
            "0102", //公司办
            "0103", //生产
            "0104", //财务
            "0109", //业务
            "0105", //人力
            "0108", //人才储备
            "0107");//总务

    /**
     * 华海部门
     */
    public final static List<String> HHAK_DEPT=Arrays.asList("HH0201","HH0203","HH020201","HH020202");
    public final static String HHAK_SELLING="HH0104";
    public final static List<String> HHAK_DEPT_NAME=Arrays.asList("安全技术咨询","管理部门","北京本部","管道项目部","东营（胜利）项目部","测绘项目部");


    /**
     * 恒远部门
     */
    public final static String HYJC="无损检测";
    public final static String ZTB="招投标保证金";
    public final static String BYJ="备用金借款";
    public final static String TAX="缴纳税费等";
    public final static List<String> PGZX=Arrays.asList("评估咨询","其中：监理公司","设计公司");
    public final static String SDSJ="工程设计";

    //设计收入部门
    public final static List<String> SDSJ_DEPT_INCOME=Arrays.asList(
            "SD02",   //北京
            "SD0301", //东营本部
            "SD05",   //总包
            "SD0401", //四川本部
            "SD0304");//评估咨询
    //设计成本部门
    public final static List<String> SDSJ_DEPT=Arrays.asList(
            "SD02",   //北京
            "SD0302", //东营生产
            "SD05",   //总包
            "SD0402", //四川生产
            "SD0305");//评估咨询

    //设计管理部门
    public final static List<String> SDSJ_OVERHEAD_DEPT=Arrays.asList(
            "SD01",   //总经办
            "SD02",   //北京
            "SD0301", //东营本部
            "SD0401", //四川本部
            "SD0304");//评估咨询



    public final static List<String> TITLE_FIRST =Arrays.asList("填报单位（公章）：","计量单位：万元");

    public final static List<String> TITLE_SECOND =Arrays.asList("序号","业务板块");

    public final static List<String> TITLE_SECOND_RIGHT =Arrays.asList("营业收入","营业总成本","经营性现金流量","企业现金流量");

    public final static String CFS_TITLE_RIGHT ="经营性现金流量";
    public final static String SALARY_TITLE_RIGHT="营业总成本-工资性支出";

    public final static List<String> TITLE_THIRD_RIGHT_MB =Arrays.asList("确认收入","挂账收入","营业总成本","毛利率");

    public final static List<String> TITLE_SIXTH_RIGHT_CFS=Arrays.asList("经营性现金流入","经营性现金流出","现金流量净额");
    public final static String COMPANY_CFS="企业现金流量";
    public final static List<String> TITLE_SIXTH_RIGHT_MAIN=Arrays.asList("经营性现金流入","期末应收账款","期末合同资产");

    public final static List<String> TITLE_SEVENTH=Arrays.asList("本年","同期","同比（%）");

    public final static List<String> TITLE_EIGHTH=Arrays.asList("本月","累计");


    /**
     * 会计账簿
     */
    public final static List<String> ACCOUNTING_BOOK = Arrays.asList(
                    "1001A2100000000008G6", //山东胜利建设监理股份有限公司
                    "1001A210000000000AGQ", //北京华海安科科技发展有限公司
                    "1001A210000000000AGT", //山东恒远检验检测有限公司
                    "1001A210000000000AGW", //北京石大东方工程设计有限公司
                    "1001A21000000001KRD3", //北京石大东方设计工程有限公司青岛分公司
                    "1001A210000000007XAO");//山东胜利建设监理股份有限公司青岛分公司

    /**
     * 会计科目编码
     */
    public final static String ACCOUNTS_RECEIVABLE="1122";//应收账款
    public final static String ACCOUNTS_RECEIVABLE_INVOICING="112201";//应收账款-开票
    public final static String ACCOUNTS_RECEIVABLE_ESTIMATED="112202";//应收账款-暂估



    /**
     * 日期
     */
    public final static String THIS_MONTH_END="2024-03-31";

}
