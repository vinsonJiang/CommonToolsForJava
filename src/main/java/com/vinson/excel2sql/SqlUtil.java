package com.vinson.excel2sql;

import java.text.MessageFormat;

/**
 * Created by JiangWeixin on 2019/2/23.
 */
public class SqlUtil {
    private static final String INSERT_SQL = "insert into {0}({1}) values({2});";
    private static final String DELETE_SQL = "delete from {0} where {1};";
    private static final String UPDATE_SQL = "update {0} set {1} where {2};";

    public static String getInsertSql(String tableName, String fieldStr, String valueStr) {
        return MessageFormat.format(INSERT_SQL, tableName, fieldStr, valueStr);
    }

    public static String getInsertSql(String tableName, String[] fields, String[] values) {

        if(fields == null || fields.length <= 0 || fields.length != values.length) {
            return null;
        }

        String fieldStr = "";
        String valueStr = "";
        for(int i = 0; i < fields.length; i++) {
            if(i == fields.length - 1) {
                fieldStr += fields[i];
                valueStr += "'" + values[i] + "'";
            } else {
                fieldStr += fields[i] + ", ";
                valueStr += "'" + values[i] + "', ";
            }
        }
        String sql = getInsertSql(tableName, fieldStr, valueStr);
        return sql;
    }

    //获取删除sql（只支持and）
    public static String getDeleteSql(String tableName, String[] fields, String[] values) {
        if(fields == null || fields.length <= 0 || fields.length != values.length) {
            return null;
        }
        String condition = "";
        for(int i = 0; i < fields.length; i++) {
            if(i == fields.length - 1) {
                condition += fields[i] + " = " + values[i];
            } else {
                condition += fields[i] + " = " + values[i] + " and ";
            }
        }
        String sql = MessageFormat.format(DELETE_SQL, tableName, condition);
        return sql;
    }

    //获取删除sql（只支持and）
    public static String getUpdateSql(String tableName, String[] fields, String[] values, String[] conditionFields, String[] conditionValues) {
        if(fields == null || fields.length <= 0 || fields.length != values.length) {
            return null;
        }
        String condition = "";
        for(int i = 0; i < conditionFields.length; i++) {
            if(i == conditionFields.length - 1) {
                condition += conditionFields[i] + " = " + conditionValues[i];
            } else {
                condition += conditionFields[i] + " = " + conditionValues[i] + " and ";
            }
        }

        String valueStrs = "";
        for(int i = 0; i < fields.length; i++) {
            if(values[i] == null || values[i].length() == 0 || "".equals(values[i].trim())) {
                continue;
            }
            valueStrs += fields[i] + " = '" + values[i] + "', ";
        }
        valueStrs = valueStrs.substring(0, valueStrs.length() - 2);

        String sql = MessageFormat.format(UPDATE_SQL,  tableName, valueStrs, condition);
        return sql;
    }
}
