package com.vinson.excelUtil;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * 找出两个Excel文件中新增的记录
 * author：jwx
 * time：2017/12/19
 * updateTime
 */
public class ExcelTool {
    //
    private static DecimalFormat df = new DecimalFormat("0");
    //格式化日期字符串
    private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
    //格式化数字
    private static DecimalFormat nf = new DecimalFormat("0");

    //是否识别单元格类型
    private final static boolean IS_REC_TYPE = true;


    public static List<List<Object>> readExcel(File file){
        if(file == null){
            return null;
        }
        if(file.getName().endsWith("xlsx")){
            return readExcel2007(file);
        }else{
            return readExcel2003(file);
        }
    }

    /**
     * 读取Excel2003
     * @param file
     * @return
     */
    public static List<List<Object>> readExcel2003(File file){
        try{
            List<List<Object>> rowList = new LinkedList<List<Object>>();
            List<Object> colList;
            HSSFWorkbook wb = new HSSFWorkbook(new FileInputStream(file));
            HSSFSheet sheet = wb.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;
            Object value;


            for(int i = sheet.getFirstRowNum(), rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows(); i++ ) {
                row = sheet.getRow(i);
                colList = new ArrayList<Object>();
                if(row == null){
                    //当读取行为空时，判断是否是最后一行
                    if(i != sheet.getPhysicalNumberOfRows()){
                        rowList.add(colList);
                    }
                    continue;
                }else{
                    rowCount++;
                }
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++){
                    cell = row.getCell(j);
                    if(cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK){
                        //当该单元格为空
                        if(j != row.getLastCellNum()){//判断是否是该行中最后一个单元格
                            colList.add("");
                        }
                        continue;
                    }

                    if(IS_REC_TYPE){
                        switch(cell.getCellType()){
                            case XSSFCell.CELL_TYPE_STRING:
                                System.out.println(i + "行" + j + " 列 is String type");
                                value = cell.getStringCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                    value = df.format(cell.getNumericCellValue());
                                } else if ("General".equals(cell.getCellStyle().getDataFormatString())) {
                                    value = nf.format(cell.getNumericCellValue());
                                } else {
                                    value = sdf.format(HSSFDateUtil.getJavaDate(cell.getNumericCellValue()));
                                }
                                System.out.println(i + "行" + j
                                        + " 列 is Number type ; DateFormt:"
                                        + value.toString());
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                System.out.println(i + "行" + j + " 列 is Boolean type");
                                value = Boolean.valueOf(cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK:
                                System.out.println(i + "行" + j + " 列 is Blank type");
                                value = "";
                                break;
                            default:
                                System.out.println(i + "行" + j + " 列 is default type");
                                value = cell.toString();
                        }
                        colList.add(value);
                    } else {
                        value = cell.toString();
                    }
                    colList.add(value);

                }
                rowList.add(colList);
            }

            return rowList;
        }catch(Exception e){
            e.printStackTrace();
            return null;
        }
    }

    /**
     * 读取Excel2007
     * @param file
     * @return
     */
    public static List<List<Object>> readExcel2007(File file){
        try{
            List<List<Object>> rowList = new LinkedList<List<Object>>();
            List<Object> colList;
            XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(file));
            XSSFSheet sheet = wb.getSheetAt(0);
            XSSFRow row;
            XSSFCell cell;
            Object value;
            for(int i = sheet.getFirstRowNum() , rowCount = 0; rowCount < sheet.getPhysicalNumberOfRows() ; i++ ){
                row = sheet.getRow(i);
                colList = new ArrayList<Object>();
                if(row == null){
                    //当读取行为空时,判断是否是最后一行
                    if(i != sheet.getPhysicalNumberOfRows()){
                        rowList.add(colList);
                    }
                    continue;
                }else{
                    rowCount++;
                }
                for( int j = row.getFirstCellNum() ; j <= row.getLastCellNum() ;j++) {
                    cell = row.getCell(j);
                    if (cell == null || cell.getCellType() == HSSFCell.CELL_TYPE_BLANK) {
                        //当该单元格为空
                        if (j != row.getLastCellNum()) {//判断是否是该行中最后一个单元格
                            colList.add("");
                        }
                        continue;
                    }
                    if(IS_REC_TYPE) {
                        switch(cell.getCellType()){
                            case XSSFCell.CELL_TYPE_STRING:
                                value = cell.getStringCellValue();
                                break;
                            case XSSFCell.CELL_TYPE_NUMERIC:
                                if ("@".equals(cell.getCellStyle().getDataFormatString())) {
                                    value = df.format(cell.getNumericCellValue());
                                } else if ("General".equals(cell.getCellStyle()
                                        .getDataFormatString())) {
                                    value = nf.format(cell.getNumericCellValue());
                                } else {
                                    value = sdf.format(HSSFDateUtil.getJavaDate(cell
                                            .getNumericCellValue()));
                                }
                                break;
                            case XSSFCell.CELL_TYPE_BOOLEAN:
                                value = Boolean.valueOf(cell.getBooleanCellValue());
                                break;
                            case XSSFCell.CELL_TYPE_BLANK:
                                value = "";
                                break;
                            default:
                                value = cell.toString();
                        }
                    } else {
                        value = cell.toString();
                    }
                    colList.add(value);
                }
                rowList.add(colList);
            }

            return rowList;
        }catch(Exception e){
            System.out.println("exception");
            return null;
        }
    }
    //写Excel2003(.xls)
    public static void writeExcel2003(List<List<Object>> result,String path){
        if(result == null){
            return;
        }
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet("sheet1");
        for(int i = 0 ;i < result.size() ; i++){
            HSSFRow row = sheet.createRow(i);
            if(result.get(i) != null){
                for(int j = 0; j < result.get(i).size() ; j++){
                    HSSFCell cell = row.createCell(j);
                    cell.setCellValue(result.get(i).get(j).toString());
                }
            }
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try
        {
            wb.write(os);
        } catch (IOException e){
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        File file = new File(path);//Excel文件生成后存储的位置。
        OutputStream fos  = null;
        try
        {
            fos = new FileOutputStream(file);
            fos.write(content);
            os.close();
            fos.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //写Excel2007(.xlsx)
    public static void writeExcel(List<List<Object>> result,String path){
        if(result == null){
            return;
        }
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("sheet1");
        for(int i = 0 ;i < result.size() ; i++){
            XSSFRow row = sheet.createRow(i);
            if(result.get(i) != null){
                for(int j = 0; j < result.get(i).size() ; j++){
                    XSSFCell cell = row.createCell(j);
                    cell.setCellValue(result.get(i).get(j).toString());
                }
            }
        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try
        {
            wb.write(os);
        } catch (IOException e){
            e.printStackTrace();
        }
        byte[] content = os.toByteArray();
        File file = new File(path);//Excel文件生成后存储的位置。
        OutputStream fos  = null;
        try
        {
            fos = new FileOutputStream(file);
            fos.write(content);
            os.close();
            fos.close();
        }catch (Exception e){
            e.printStackTrace();
        }
    }

    //newFile中相对于在oldFile中没有的记录
    /**
     * 查找差异道具（包括新增和修改）
     * @param oldFileName
     * @param newFileName
     * @param rstFileName
     */
    public static void getDifByRecord(String oldFileName, String newFileName, String rstFileName){
        File file1 = new File(oldFileName);
        File file2 = new File(newFileName);

        List<List<Object>> excel1 = readExcel(file1);
        List<List<Object>> excel2 = readExcel(file2);
        List<List<Object>> result = new ArrayList<List<Object>>();

        Map<String, List<Object>> map = new HashMap<String, List<Object>>();

        StringBuilder sb = new StringBuilder();
        System.out.println(excel2.size());

        for(int i = 0 ;i < excel2.size() ;i++){
            sb.setLength(0);
            for(int j = 0; j < excel2.get(i).size(); j++){
                sb.append(excel2.get(i).get(j).hashCode());
            }
            String hash = sb.toString();//getHashKey(sb.toString());
            map.put(hash, excel2.get(i));

        }

        for(int i = 0 ;i < excel1.size() ;i++){
            sb.setLength(0);
            for(int j = 0; j < excel1.get(i).size(); j++){
                sb.append(excel1.get(i).get(j).hashCode());
            }
            String hash = sb.toString();//getHashKey(sb.toString());
            if(map.containsKey(hash)){
                map.remove(hash);
            }
        }

        for(Map.Entry<String, List<Object>> item : map.entrySet()){
            result.add(item.getValue());
        }

        writeExcel(result, rstFileName);
    }

    //newFile中相对于oldFile中新增的id
    /**
     * 查找增量道具（新增的id）
     * @param oldFileName
     * @param newFileName
     * @param rstFileName
     */
    public static void getDifById(String oldFileName, String newFileName, String rstFileName){
        File file1 = new File(oldFileName);
        File file2 = new File(newFileName);

        List<List<Object>> excel1 = readExcel(file1);
        List<List<Object>> excel2 = readExcel(file2);
        List<List<Object>> result = new ArrayList<List<Object>>();

        Map<Integer, List<Object>> map = new HashMap<Integer, List<Object>>();


        for(int i = 0 ;i < excel2.size() ;i++){
            int hash = excel2.get(i).get(0).hashCode();
            map.put(hash, excel2.get(i));
        }

        for(int i = 0 ;i < excel1.size() ;i++){
            int hash = excel1.get(i).get(0).hashCode();
            if(map.containsKey(hash)){
                map.remove(hash);
            }
        }

        for(Map.Entry<Integer, List<Object>> item : map.entrySet()){
            result.add(item.getValue());
        }


        writeExcel(result, rstFileName);

    }

    //去掉差量道具
    public static void removeDifItem(String oldFileName, String newFileName, String rstFileName){
        File file1 = new File(oldFileName);
        File file2 = new File(newFileName);

        List<List<Object>> excel1 = readExcel(file1);
        List<List<Object>> excel2 = readExcel(file2);
        List<List<Object>> result = new ArrayList<List<Object>>();

        Map<Integer, List<Object>> map = new HashMap<Integer, List<Object>>();

        Set<Integer> set = new HashSet<Integer>();

        for(int i = 0 ;i < excel1.size() ;i++){
            String itemId = (String)excel1.get(i).get(0);
            int itemIdInt = Integer.parseInt(itemId);
            int infoIdInt = itemIdInt + 100000000;
            int hash = (infoIdInt + "").hashCode();
            set.add(hash);
        }

        Iterator<List<Object>> it = excel2.iterator();
        while(it.hasNext()){
            List<Object> objs = it.next();
            int hash = objs.get(0).hashCode();

            if(set.contains(hash)){
                it.remove();
//                System.out.println("HHH");
            }
        }

//        for(Map.Entry<Integer, List<Object>> item : map.entrySet()){
//            result.add(item.getValue());
//        }

        writeExcel(excel2, rstFileName);
    }

    //分组求和
    public static void getSumByGroup(String oldFileName, String rstFileName) {
        getSumByGroup(oldFileName, rstFileName, 0, 1);
    }
    public static void getSumByGroup(String oldFileName, String rstFileName, int groupIndex, int targetIndex){
        File file1 = new File(oldFileName);
        List<List<Object>> excel1 = readExcel(file1);
        Map<String, List<Object>> map = new HashMap<String, List<Object>>();

        for(List<Object> item : excel1){
            String acountId = (String)item.get(groupIndex);
            int num = Integer.parseInt(item.get(targetIndex).toString());
            if(map.containsKey(acountId)){
                List<Object> list = map.get(acountId);
                //int num = (Integer) list.get(1);
                int oldNum = Integer.parseInt(list.get(targetIndex).toString());
                list.set(targetIndex, oldNum + num);
            }else {
                map.put(acountId, item);
            }
        }

        //List<List<Object>> excel2 = (List<List<Object>>)map.values();
        List<List<Object>> excel2 = new ArrayList<List<Object>>();
        for(Map.Entry<String, List<Object>> item : map.entrySet()){
            excel2.add(item.getValue());
        }

        //处理Excel2003最大行数为65535行，超过每个文件放60000行
//        int size = excel2.size();

//        if(size > 65535){
//            int remain = size;
//            int no = 0;
//            while(remain > 0) {
//                writeExcel(excel2.subList(size - remain, Math.min(size - remain + 60000, size)), rstFileName + (no++) + ".xls");
//                remain -= 60000;
//            }
//        }else {
//            writeExcel(excel2, rstFileName + ".xls");
//        }
        writeExcel(excel2, rstFileName);
    }


    public static void main(String[] args) {
        String oldFileName = "F:/test/海外.xlsx";
        String rstFileName = "F:/test/海外result.xlsx";

        getSumByGroup(oldFileName, rstFileName);
    }

}