package com.vinson;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;

/**
 * @Description:
 * @author: jiangweixin
 * @date: 2019/2/4
 */
public class FileUtil {
    public static void outputToFile(File file, String content) throws Exception {
        FileOutputStream outStream = null;
        BufferedOutputStream buffer = null;
        outStream = new FileOutputStream(file);
        buffer = new BufferedOutputStream(outStream);
        buffer.write(content.getBytes());
        buffer.close();
    }
}
