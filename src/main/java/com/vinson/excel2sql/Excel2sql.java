package com.vinson.excel2sql;

import com.vinson.FileUtil;
import com.vinson.excelUtil.ExcelTool;

import java.io.File;
import java.util.Arrays;
import java.util.List;

/**
 * Created by JiangWeixin on 2018/3/15.
 */
public class Excel2sql {

    /** 一行生成的sql数 */
    private static final int NUM = 11;
    /** 开始行（0开始算） */
    private static final int START_COL = 1;
    /** 插入或更新 */
    private static final boolean IS_INSERT = true;
    /** 是否在excel中定义字段 */
    private static final boolean HAS_FIELDS = true;
    /** 字段 */
    private static final int FIELDS_NUM = 7;
    private static final String[] FIELDS = {"info_id", "`param2`"};
    /** 表名 */
    private static final String TABLE_NAME = "t_table_name";
    /** 主键个数 */
    private static final int KEY_NUM = 2;

    private static final String INSERT_SQL = "insert into {0}({1}) values({2});";
    private static final String DELETE_SQL = "delete from {0} where {1};";
    private static final String UPDATE_SQL = "update {0} {1} where {1}";

    public static void main(String[] args) {
        String fileName = "";
        String outputFileName = "";

        new Excel2sql().getResult(fileName, outputFileName, IS_INSERT);
    }

    /**
     * Excel生成insert，delete语句
     */
    public void getResult(String srcFile, String desFile, boolean isInsert) {

        File file1 = new File(srcFile);
        File file2 = new File(desFile);
        List<List<Object>> excel1 = ExcelTool.readExcel(file1);

        StringBuilder insertSb = new StringBuilder();
        StringBuilder deleteSb = new StringBuilder();
        StringBuilder sqlSb = new StringBuilder();

        String[] fields = new String[FIELDS_NUM];
        if(!HAS_FIELDS) {
            fields = FIELDS;
        }

        int lineNum = 0;
        for(List<Object> line : excel1) {
            lineNum++;
            if(lineNum <= START_COL) {
                continue;
            } else if(HAS_FIELDS && lineNum == START_COL + 1) {
                for(int j = 0; j < FIELDS_NUM; j++) {
                    fields[j] = "`" + line.get(j) + "`";
                }
                continue;
            }

            for(int i = 0; i < NUM; i++) {
                String[] values = new String[FIELDS_NUM];
                for(int j = 0; j <FIELDS_NUM; j++) {
                    int index = i * FIELDS_NUM + j;
                    String temp;
                    if(line.size() < index || line.get(index) == null) {
                        temp = "";
                    } else {
                        temp = line.get(index).toString();
                    }
                    values[j] = temp.trim().replace("\\'", "'").replace("'", "\\'")
                            .replace("\n", "\\n").replace("\t", "\\t");
                }

                if(isInsert) {
                    String insertSql = SqlUtil.getInsertSql(TABLE_NAME, fields, values);
                    String deleteSql = SqlUtil.getDeleteSql(TABLE_NAME, Arrays.copyOfRange(fields, 0, KEY_NUM), Arrays.copyOfRange(values, 0, KEY_NUM));
                    System.out.println(insertSql);
                    sqlSb.append(deleteSql + System.lineSeparator());
                    sqlSb.append(insertSql + System.lineSeparator());
                } else {
                    String updateSql = SqlUtil.getUpdateSql(TABLE_NAME, Arrays.copyOfRange(fields, KEY_NUM, fields.length), Arrays.copyOfRange(values, KEY_NUM, fields.length),
                            Arrays.copyOfRange(fields, 0, KEY_NUM), Arrays.copyOfRange(values, 0, KEY_NUM));
                    System.out.println(updateSql);
                    sqlSb.append(updateSql + System.lineSeparator());
                }
            }
        }

        try {
            FileUtil.outputToFile(file2, sqlSb.toString());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }




}
