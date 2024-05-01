package Lab2;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;


import java.io.FileInputStream;
import java.io.IOException;

public class ExcelParser {

    public static ArrayList<Sample> parse(String fileName, int sheetNumber) throws IOException {
        FileInputStream file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheetAt(sheetNumber+1);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            throw  new IOException("Sheet is empty");
        }
        Map<Integer, List<Double>> data = new HashMap<>();
        int sampleNumber=0;
        for (Row row : sheet) {
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    int columnIndex = cell.getColumnIndex();
                    if (!data.containsKey(columnIndex)) {
                        data.put(sampleNumber, new ArrayList<>());
                        sampleNumber++;
                    }
                    data.get(columnIndex).add(cell.getNumericCellValue());
                }
            }
        }
        ArrayList<Sample> result=new ArrayList<>();
        data.forEach((integer, doubles) -> result.add(new Sample(doubles,integer)) );
        return result;
    }
}


