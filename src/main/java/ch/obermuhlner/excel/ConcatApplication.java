package ch.obermuhlner.excel;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class ConcatApplication {
    public static void main(String[] args) {
        String targetFilename = args[0];

        XSSFWorkbook targetWorkbook = new XSSFWorkbook();
        Map<String, XSSFSheet> targetSheetMap = new HashMap<>();
        Map<String, Integer> targetLastRowIndexMap = new HashMap<>();

        CellCopyPolicy cellCopyPolicy = new CellCopyPolicy.Builder().cellValue(true).cellStyle(false).build();

        for (int i = 1; i < args.length; i++) {
            String sourceFilename = args[i];
            try (XSSFWorkbook sourceWorkbook = new XSSFWorkbook(new File(sourceFilename))) {
                int sourceSheetCount = sourceWorkbook.getNumberOfSheets();
                for (int sourceSheetIndex = 0; sourceSheetIndex < sourceSheetCount; sourceSheetIndex++) {
                    XSSFSheet sourceSheet = sourceWorkbook.getSheetAt(sourceSheetIndex);
                    String sheetName = sourceSheet.getSheetName();
                    XSSFSheet targetSheet = targetSheetMap.get(sheetName);
                    int targetRowIndex = targetLastRowIndexMap.computeIfAbsent(sheetName, k -> 0);
                    if (targetSheet == null) {
                        targetSheet = targetWorkbook.createSheet(sheetName);
                        targetSheetMap.put(sheetName, targetSheet);
                    }

                    int firstSourceRow = sourceSheet.getFirstRowNum();
                    int lastSourceRow = sourceSheet.getLastRowNum();
                    for (int sourceRowIndex = firstSourceRow; sourceRowIndex <= lastSourceRow; sourceRowIndex++) {
                        XSSFRow sourceRow = sourceSheet.getRow(sourceRowIndex);
                        if (sourceRow != null) {
                            XSSFRow targetRow = targetSheet.createRow(targetRowIndex++);

                            int firstSourceCell = sourceRow.getFirstCellNum();
                            int lastSourceCell = sourceRow.getLastCellNum();
                            for (int sourceCellIndex = firstSourceCell; sourceCellIndex <= lastSourceCell; sourceCellIndex++) {
                                XSSFCell sourceCell = sourceRow.getCell(sourceCellIndex);
                                if (sourceCell != null) {
                                    XSSFCell targetCell = targetRow.createCell(sourceCellIndex);
                                    targetCell.copyCellFrom(sourceCell, cellCopyPolicy);
                                }
                            }
                        }
                    }

                    targetLastRowIndexMap.put(sheetName, targetRowIndex);
                }
            } catch (InvalidFormatException | IOException e) {
                e.printStackTrace();
            }
        }

        try (OutputStream out = new BufferedOutputStream(new FileOutputStream(targetFilename))) {
            targetWorkbook.write(out);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
