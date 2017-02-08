package net.therap.main;

import net.therap.util.WorkbookSplitter;

/**
 * @author shakhawat.hossain
 * @since 2/8/17
 */
public class ExcelUtilityApp {

    private static final String filePath = "/home/shakhawat.hossain/upload/file_name.xlsx";
    private static final int maxRowPerExcel = 300;

    public static void main(String[] args) {
        WorkbookSplitter workbookSplitter = new WorkbookSplitter(filePath, maxRowPerExcel);
        workbookSplitter.splitWorkbook();
    }
}
