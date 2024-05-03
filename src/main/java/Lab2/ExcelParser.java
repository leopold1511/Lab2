package Lab2;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;


import java.io.FileInputStream;
import java.io.IOException;

public class ExcelParser {

    public static ArrayList<Sample> parse(String fileName, String sheetName) throws IOException {
        FileInputStream file = new FileInputStream(fileName);
        XSSFWorkbook workbook = new XSSFWorkbook(file);
        XSSFSheet sheet = workbook.getSheet(sheetName);
        if (sheet.getPhysicalNumberOfRows() == 0) {
            throw  new IOException("Sheet is empty");
        }
        Map<Integer, List<Double>> data = new HashMap<>();
        for (Row row : sheet) {
            int sampleNumber=1;
            for (Cell cell : row) {
                if (cell.getCellType() == CellType.NUMERIC) {
                    if (!data.containsKey(sampleNumber)) {
                        data.put(sampleNumber, new ArrayList<>());
                    }
                    data.get(sampleNumber).add(cell.getNumericCellValue());
                }
                sampleNumber++;
            }
        }
        ArrayList<Sample> result=new ArrayList<>();
        data.forEach((integer, doubles) -> result.add(new Sample(doubles,integer)) );
        return result;
    }
}


