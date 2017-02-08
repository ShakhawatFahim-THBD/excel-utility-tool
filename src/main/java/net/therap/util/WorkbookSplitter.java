package net.therap.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import static net.therap.util.StringUtil.isNotEmpty;

/**
 * @author shakhawat.hossain
 * @since 2/8/17
 */
public class WorkbookSplitter {

    private String filePath;
    private int maxRowPerFile;

    public WorkbookSplitter(String fileName, final int maxRowPerFile) {
        this.filePath = fileName;
        this.maxRowPerFile = maxRowPerFile;
    }

    public void splitWorkbook() {
        List<Workbook> workbooks = getWorkbooks();

        System.out.println("========== Completed Processing Input Excel From: " + filePath + " ==================");
        System.out.println("========== Total File to be created : " + workbooks.size() + " ==================");

        writeWorkBooks(workbooks);
    }

    public List<Workbook> getWorkbooks() {
        Workbook inputWorkBook = null;

        try {
            inputWorkBook = WorkbookFactory.create(new File(filePath));
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }

        if (inputWorkBook == null || inputWorkBook.getSheetAt(0) == null || inputWorkBook.getNumberOfSheets() == 0) {
            return Collections.emptyList();
        }

        Sheet inputSheet = inputWorkBook.getSheetAt(0);

        System.out.println("========== Total Number of Row to be processed (with blank rows): "
                + inputSheet.getPhysicalNumberOfRows() + " ==================");

        Row headerRow = null;
        int totalProcessedRowCount = 0;

        Iterator<Row> rowIterator = inputSheet.iterator();
        if (rowIterator.hasNext()) {
            headerRow = rowIterator.next();
            totalProcessedRowCount++;
        }

        assert headerRow != null;

        int headerCount = findLastNonEmptyColumnIndex(headerRow);

        List<Workbook> workbooks = new ArrayList<>();
        Workbook workbook = new HSSFWorkbook();
        Sheet splittedInputSheet = workbook.createSheet();
        Row newHeaderRow = splittedInputSheet.createRow(0);
        copyRowValues(headerRow, newHeaderRow, headerCount);

        int rowCount = 0;

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();

            if (!isEmpty(row)) {
                Row newRow = splittedInputSheet.createRow(splittedInputSheet.getLastRowNum() + 1);
                copyRowValues(row, newRow, headerCount);
                totalProcessedRowCount++;
                rowCount++;
            }

            if (rowCount == maxRowPerFile) {
                workbooks.add(workbook);

                workbook = new HSSFWorkbook();
                splittedInputSheet = workbook.createSheet();
                newHeaderRow = splittedInputSheet.createRow(0);
                copyRowValues(headerRow, newHeaderRow, headerCount);
                rowCount = 0;
            }
        }

        if (workbook.getSheetAt(0).getPhysicalNumberOfRows() > 0) {
            workbooks.add(workbook);
        }

        System.out.println("========== Total Number of Row to be processed (without blank rows): "
                + totalProcessedRowCount + " ==================");

        return workbooks;
    }

    private int findLastNonEmptyColumnIndex(Row row) {
        if (row == null) {
            return 0;
        }

        int lastNonEmptyIndex = 0;
        int index = 0;
        Iterator<Cell> cellIterator = row.cellIterator();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();

            if (cell.getCellTypeEnum() != CellType.BLANK) {
                lastNonEmptyIndex = index;
            }

            index++;
        }

        return lastNonEmptyIndex;
    }

    private void copyRowValues(Row existingRow, Row newRow, int headerCount) {
        int index = 0;
        Iterator<Cell> cellIterator = existingRow.cellIterator();

        while (cellIterator.hasNext()) {
            Cell existingCell = cellIterator.next();
            CellType cellType = existingCell.getCellTypeEnum();

            if (existingCell.getColumnIndex() >= headerCount) {
                break;
            }

            Cell newCell = newRow.createCell(index++);

            newCell.setCellType(cellType);

            switch (cellType) {
                case NUMERIC:
                    newCell.setCellValue(existingCell.getNumericCellValue());
                    break;
                case BLANK:
                case STRING:
                    newCell.setCellValue(existingCell.getStringCellValue());
                    break;
                case BOOLEAN:
                    newCell.setCellValue(existingCell.getBooleanCellValue());
                    break;
                default:
                    throw new IllegalStateException(cellType + " has not been handled");
            }
        }
    }

    private boolean isEmpty(Row row) {
        if (row == null) {
            return true;
        }

        Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();

            if (cell != null && isNotEmpty(cell.getStringCellValue())) {
                return false;
            }
        }

        return true;
    }

    private void writeWorkBooks(List<Workbook> workbooks) {
        int totalRowWritten = 0;
        int extensionIndex = filePath.lastIndexOf(".");

        String fileNameWitoutPrefix = filePath.substring(0, extensionIndex);
        String extension = filePath.substring(extensionIndex);

        FileOutputStream out = null;
        try {
            for (int i = 0; i < workbooks.size(); i++) {
                String newFileName = fileNameWitoutPrefix + ("_" + (i + 1)) + extension;

                out = new FileOutputStream(new File(newFileName));
                workbooks.get(i).write(out);

                int rowCount = workbooks.get(i).getSheetAt(0).getPhysicalNumberOfRows();
                totalRowWritten += rowCount;

                System.out.println("========== Successfully wrote file " + (i + 1) + ", #row: " + rowCount + " ===================");
            }
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        System.out.println("========== Successfully wrote: " + totalRowWritten + " rows ===================");
    }
}