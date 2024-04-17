package Lab2;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ExcelCreator {

    public void writeSamplesToExcel(List<Sample> samples, String filePath) {
        try (Workbook workbook = new XSSFWorkbook()) {

            Sheet sheet1 = workbook.createSheet("Лист 1");
            writeSampleResults(sheet1, samples);

            Sheet sheet2 = workbook.createSheet("Матрица ковариации");
            writeCovarianceMatrix(sheet2, samples);

            try (FileOutputStream fileOut = new FileOutputStream(filePath)) {
                workbook.write(fileOut);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private void writeSampleResults(Sheet sheet, List<Sample> samples) {
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Номер выборки");
        String[] headers = {"Геометрическое среднее", "Арифметическое среднее", "Стандартное отклонение",
                "Размах", "Количество элементов", "Коэффициент вариации", "Нижняя граница доверительного интервала",
                "Верхняя граница доверительного интервала", "Дисперсия", "Минимальное значение", "Максимальное значение"};
        for (int i = 0; i < headers.length; i++) {
            headerRow.createCell(i + 1).setCellValue(headers[i]);
        }

        int rowNum = 1;
        for (Sample sample : samples) {
            Row row = sheet.createRow(rowNum++);
            row.createCell(0).setCellValue(sample.id);
            double[] results = sample.getResult();
            for (int i = 0; i < results.length; i++) {
                row.createCell(i + 1).setCellValue(results[i]);
            }
        }
    }

    private void writeCovarianceMatrix(Sheet sheet, List<Sample> samples) {

        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Номер выборки");
        for (int i = 0; i < samples.size(); i++) {
            headerRow.createCell(i + 1).setCellValue("Выборка " + samples.get(i).id);
        }

        for (int i = 0; i < samples.size(); i++) {
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue("Выборка " + samples.get(i).id);
            for (int j = 0; j < samples.size(); j++) {
                row.createCell(j + 1).setCellValue(samples.get(i).calculateCovariance(samples.get(j).sample));
            }
        }
    }
}
