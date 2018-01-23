package com.vinson.fileLineCount;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.LinkedList;
import java.util.Queue;

/**
 * Created by JiangWeixin on 2018/1/23.
 */
public class FileLineCount {
    public static void main(String[] args) {
        String filepath = "F:\\Project\\IdeaProjects";

        int count = getResult(filepath);
        System.out.println("代码总行数：" + count);
    }

    /**
     * 获取目录下所有文件代码总行数
     * @param filepath
     * @return
     */
    public static int getResult(String filepath) {

        Queue<File> fileQueue = new LinkedList<File>();

        int count = 0;

        File file = new File(filepath);
        File[] fileList = file.listFiles();

        for(File item : fileList) {
            fileQueue.add(item);
        }
        while(!fileQueue.isEmpty()) {
            File temp = fileQueue.poll();
            if (temp.isFile()) {//判断是否为文件
                count += getFileLineCounts(temp);
            } else if (temp.isDirectory()) {
                fileList = temp.listFiles();
                for(File item : fileList) {
                    fileQueue.add(item);
                }
            }
        }
        return count;
    }

    /**
     * 计算文件代码总行数
     * @param fileName
     * @return
     */
    public static int getFileLineCounts(String fileName) {
        File file = new File(fileName);
        return getFileLineCounts(file);
    }

    public static int getFileLineCounts(File file) {
        int cnt = 0;
        InputStream is = null;
        try {
            is = new BufferedInputStream(new FileInputStream(file));
            byte[] c = new byte[1024];
            int readChars = 0;
            while ((readChars = is.read(c)) != -1) {
                for (int i = 0; i < readChars; ++i) {
                    if (c[i] == '\n') {
                        ++cnt;
                    }
                }
            }
        } catch (Exception ex) {
            cnt = -1;
            ex.printStackTrace();
        } finally {
            try {
                is.close();
            } catch (Exception ex) {
                ex.printStackTrace();
            }
        }
        return cnt;
    }
}
