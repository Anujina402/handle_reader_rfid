package com.pda.uhf_g.util;


import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelUtil {
    public static WritableFont arial14font = null;

    public static WritableCellFormat arial14format = null;
    public static WritableFont arial10font = null;
    public static WritableCellFormat arial10format = null;
    public static WritableFont arial12font = null;
    public static WritableCellFormat arial12format = null;

    public final static String UTF8_ENCODING = "UTF-8";
    public final static String GBK_ENCODING = "GBK";


    /**

     */
    public static void format() {
        try {
            arial14font = new WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD);
            arial14font.setColour(Colour.LIGHT_BLUE);
            arial14format = new WritableCellFormat(arial14font);
            arial14format.setAlignment(jxl.format.Alignment.CENTRE);
            arial14format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            arial14format.setBackground(Colour.VERY_LIGHT_YELLOW);

            arial10font = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            arial10format = new WritableCellFormat(arial10font);
            arial10format.setAlignment(jxl.format.Alignment.CENTRE);
            arial10format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
            arial10format.setBackground(Colour.GRAY_25);

            arial12font = new WritableFont(WritableFont.ARIAL, 10);
            arial12format = new WritableCellFormat(arial12font);
            arial10format.setAlignment(jxl.format.Alignment.CENTRE);//对齐格式
            arial12format.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN); //设置边框

        } catch (WriteException e) {
            e.printStackTrace();
        }
    }

    /**

     *
     * @param fileName
     * @param colName
     */
    public static File initExcel(File directory, String fileName, String[] colName) throws IOException {
        format();
        WritableWorkbook workbook = null;
        File file = new File(directory, fileName);
        try {
            if (!directory.exists() && !directory.mkdirs()) {
                throw new IOException("Failed to create directory: " + directory.getAbsolutePath());
            }
            if (file.exists() && !file.delete()) {
                throw new IOException("Failed to overwrite file: " + file.getAbsolutePath());
            }
            workbook = Workbook.createWorkbook(file);
            String sheetName = fileName;
            int extensionIndex = sheetName.lastIndexOf('.');
            if (extensionIndex > 0) {
                sheetName = sheetName.substring(0, extensionIndex);
            }
            if (sheetName.isEmpty()) {
                sheetName = "Sheet1";
            }
            WritableSheet sheet = workbook.createSheet(sheetName, 0);
            //
            sheet.addCell((WritableCell) new Label(0, 0, fileName, arial14format));
            for (int col = 0; col < colName.length; col++) {
                sheet.addCell(new Label(col, 0, colName[col], arial10format));
            }
            sheet.setRowView(0, 340); //
            workbook.write();
        } catch (Exception e) {
            throw new IOException("Failed to initialize excel file", e);
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    // Ignore close exceptions to preserve the original failure cause.
                }
            }
        }
        return file;
    }

    @SuppressWarnings("unchecked")
    public static <T> void writeObjListToExcel(ArrayList<ArrayList<String>> objList, File file) throws IOException {
        if (objList == null || objList.isEmpty()) {
            return;
        }
        WritableWorkbook writebook = null;
        InputStream in = null;
        try {
            WorkbookSettings setEncode = new WorkbookSettings();
            setEncode.setEncoding(UTF8_ENCODING);
            in = new FileInputStream(file);
            Workbook workbook = Workbook.getWorkbook(in);
            writebook = Workbook.createWorkbook(file, workbook);
            WritableSheet sheet = writebook.getSheet(0);

            for (int j = 0; j < objList.size(); j++) {
                ArrayList<String> list = (ArrayList<String>) objList.get(j);
                for (int i = 0; i < list.size(); i++) {
                    sheet.addCell(new Label(i, j + 1, list.get(i), arial12format));
                    if (list.get(i).length() <= 5) {
                        sheet.setColumnView(i, list.get(i).length() + 8); //
                    } else {
                        sheet.setColumnView(i, list.get(i).length() + 5); //
                    }
                }
                sheet.setRowView(j + 1, 350); //
            }

            writebook.write();
        } catch (Exception e) {
            throw new IOException("Failed to write excel file", e);
        } finally {
            if (writebook != null) {
                try {
                    writebook.close();
                } catch (Exception e) {
                    // Ignore close exceptions to preserve the original failure cause.
                }

            }
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e) {
                    // Ignore close exceptions to preserve the original failure cause.
                }
            }
        }
    }
}





