package Lab2;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class ExcelParser {

    public static Map<Integer, List<Double>> parse(String fileName, int sheetNumber) throws IOException {
        FileInputStream file = new FileInputStream(new File(fileName));
        Workbook workbook = new XSSFWorkbook(file);
        Sheet sheet = workbook.getSheetAt(sheetNumber);
        Map<Integer, List<Double>> data = new HashMap<>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    int columnIndex = cell.getColumnIndex();
                    if (!data.containsKey(columnIndex)) {
                        data.put(columnIndex, new ArrayList<>());
                    }
                    data.get(columnIndex).add(cell.getNumericCellValue());
                }
            }
        }
        return data;
    }
}


