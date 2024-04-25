package cn.sljl.mapper;

import java.util.List;

/**
 * 查询金额
 * @author wangeqiu
 * @version 1.0
 * @date 2024/3/16 8:57
 */
public class CommonSql {

    /**
     * 查询一级科目【贷方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     */
    public String creditAmount(String startDate, String endDate, String accountBook, String ledgerAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALCREDITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code like '"+ledgerAccount+"%')\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询一级科目【借方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     */
    public String debitAmount(String startDate,String endDate,String accountBook,String ledgerAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALDEBITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code like '"+ledgerAccount+"%')\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     *查询二级科目【贷方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     */
    public String creditAmount(String startDate, String endDate, String accountBook, String ledgerAccount,String specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALCREDITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code like '"+specificAccount+"%')\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询二级科目【借方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @return
     */
    public String debitAmount(String startDate,String endDate,String accountBook,String ledgerAccount,String specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALDEBITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code like '"+specificAccount+"%')\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询单一末级科目【贷方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @param specificAccount
     * @return
     */
    public String singleCreditAmount(String startDate, String endDate, String accountBook, String ledgerAccount, String specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALCREDITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code in ('"+specificAccount+"'))\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询单一末级科目【借方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @param specificAccount
     * @return
     */
    public String singleDebitAmount(String startDate,String endDate,String accountBook,String ledgerAccount,String specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALDEBITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code in ('"+specificAccount+"'))\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询四个末级科目【借方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @param specificAccount
     * @return
     */
    public String quadraDebitAmount(String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALDEBITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code in ('"+specificAccount.get(0)+"','"+specificAccount.get(1)+"','"+specificAccount.get(2)+"','"+specificAccount.get(3)+"'))\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

    /**
     * 查询三个末级科目【借方】
     * @param startDate
     * @param endDate
     * @param accountBook
     * @param ledgerAccount
     * @param specificAccount
     * @return
     */
    public String tripleDebitAmount(String startDate, String endDate, String accountBook, String ledgerAccount, List<String> specificAccount){
        String sql="select accode, acname,dept_code,dept_name,sum(nvl(cmonth,0)) amount from (\n" +
                "    select bd_account.code accode,bd_accasoa.name acname,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.CODE else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.CODE else '' end) dept_code,\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE1,1,20)='0001Z0100000000005CS' then dept1.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE2,1,20)='0001Z0100000000005CS' then dept2.name else '' end) ||\n" +
                "       (case when substr(gl_freevalue.TYPEVALUE3,1,20)='0001Z0100000000005CS' then dept3.name else '' end) dept_name,\n" +
                "       (case when substr(bd_account.code,1,4) = '"+ledgerAccount+"' then GL_DETAIL.LOCALDEBITAMOUNT else 0 end) cmonth\n" +
                "    from bd_account\n" +
                "        inner join bd_accasoa on bd_accasoa.pk_account=bd_account.PK_ACCOUNT\n" +
                "        left join gl_detail on bd_accasoa.pk_accasoa = gl_detail.PK_ACCASOA\n" +
                "        left join gl_voucher on gl_voucher.pk_voucher=gl_detail.pk_voucher\n" +
                "        left join gl_freevalue on gl_freevalue.FREEVALUEID=gl_detail.ASSID\n" +
                "        left join org_dept dept1 on dept1.pk_dept=substr(gl_freevalue.TYPEVALUE1,21,20)\n" +
                "        left join org_dept dept2 on dept2.pk_dept=substr(gl_freevalue.TYPEVALUE2,21,20)\n" +
                "        left join org_dept dept3 on dept3.pk_dept=substr(gl_freevalue.TYPEVALUE3,21,20)\n" +
                "        left join org_accountingbook on org_accountingbook.PK_ACCOUNTINGBOOK=gl_voucher.PK_ACCOUNTINGBOOK\n" +
                "    where (bd_account.code in ('"+specificAccount.get(0)+"','"+specificAccount.get(1)+"','"+specificAccount.get(2)+"'))\n" +
                "      and nvl(gl_detail.dr,0)=0\n" +
                "      and nvl(gl_voucher.dr,0)=0\n" +
                "      and nvl(gl_freevalue.dr,0)=0\n" +
                "      and org_accountingbook.pk_accountingbook='"+accountBook+"'\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD >=  substr( '"+startDate+"',1,7)\n" +
                "      and gl_voucher.year || '-' || gl_voucher.PERIOD <=  substr( '"+endDate+"',1,7)\n" +
                "    ) tt\n" +
                "group by dept_name,dept_code,accode,acname\n" +
                "order by dept_code,accode";
        return sql;
    }

}
