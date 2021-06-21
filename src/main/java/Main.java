import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class Main {

    public static void main(String[] args) {
        try {
            ArrayList<Model> models = readFromExcel();
            String excelFilePath = "data.xls";
            writeExcel(models, excelFilePath);

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void writeExcel(List<Model> models, String excelFilePath) throws IOException {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();

        int rowCount = 0;
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("DATETIME");
        cell = row.createCell(1);
        cell.setCellValue("UPPER BAND");
        cell = row.createCell(2);
        cell.setCellValue("LOWER BAND");
        cell = row.createCell(3);
        cell.setCellValue("HIGH DATETIME");
        cell = row.createCell(4);
        cell.setCellValue("HIGH");
        cell = row.createCell(5);
        cell.setCellValue("LOW DATETIME");
        cell = row.createCell(6);
        cell.setCellValue("LOW");
        for (Model model : models) {
            row = sheet.createRow(++rowCount);
            writeModel(model, row, workbook);
        }
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);
        sheet.autoSizeColumn(4);
        sheet.autoSizeColumn(5);
        sheet.autoSizeColumn(6);
        try (FileOutputStream outputStream = new FileOutputStream(excelFilePath)) {
            workbook.write(outputStream);
        }
    }

    private static void writeModel(Model model, Row row, Workbook workbook) {
        DataFormat format = workbook.createDataFormat();
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.setDataFormat(format.getFormat("yyyy-MM-dd hh:mm"));
        Cell cell = row.createCell(0);
        cell.setCellStyle(dateStyle);
        cell.setCellValue(model.getDateTime());


        cell = row.createCell(1);
        cell.setCellValue(model.getUpperBand());

        cell = row.createCell(2);
        cell.setCellValue(model.getLowerBand());

        cell = row.createCell(3);
        if (model.getDateTimeForHigh() != null) {
            cell.setCellStyle(dateStyle);
            cell.setCellValue(model.getDateTimeForHigh());
        }

        cell = row.createCell(4);
        if (model.getHigh() != 0.0) {
            cell.setCellValue(model.getHigh());
        }

        cell = row.createCell(5);
        if (model.getDateTimeForLow() != null) {
            cell.setCellStyle(dateStyle);
            cell.setCellValue(model.getDateTimeForLow());
        }
        cell = row.createCell(6);
        if (model.getLow() != 0.0) {
            cell.setCellValue(model.getLow());
        }
    }

    public static ArrayList<Model> readFromExcel() throws IOException {
        FileInputStream fis;
        Properties property = new Properties();

        fis = new FileInputStream("src/main/resources/application.properties");
        property.load(fis);

        String hourlyDataFile = property.getProperty("hourFile");
        String minuteDataFile = property.getProperty("minuteFile");

        FileInputStream hourFile = new FileInputStream(new File(hourlyDataFile));
        Workbook workbookForHour = new XSSFWorkbook(hourFile);
        Sheet sheetForHour = workbookForHour.getSheetAt(0);
        FileInputStream minuteFile = new FileInputStream(new File(minuteDataFile));
        Workbook workbookForMinute = new XSSFWorkbook(minuteFile);
        Sheet sheetForMinute = workbookForMinute.getSheetAt(0);
        double high;
        double low;
        double upperBand;
        double lowerBand;
        Date dateTimeOfHourData;
        ArrayList<Model> models = new ArrayList<>();
        int i = 3;
        System.out.println(sheetForHour.getLastRowNum());
        while (i < sheetForHour.getLastRowNum()) {
            try {
                lowerBand = sheetForHour.getRow(i).getCell(11).getNumericCellValue();
                upperBand = sheetForHour.getRow(i).getCell(9).getNumericCellValue();
                low = sheetForHour.getRow(i).getCell(6).getNumericCellValue();
                high = sheetForHour.getRow(i).getCell(5).getNumericCellValue();
                dateTimeOfHourData = sheetForHour.getRow(i).getCell(3).getDateCellValue();
                Model model = new Model();
                if (high > upperBand & low < lowerBand) {
                    model.setDateTime(dateTimeOfHourData);
                    model.setLowerBand(lowerBand);
                    model.setUpperBand(upperBand);
                    findMinuteData(model, sheetForMinute, models);
                }
                i++;
            } catch (NullPointerException e) {
                if (i != sheetForHour.getLastRowNum()) {
                    i++;
                } else {
                    break;
                }
            }

        }
        return models;
    }

    private static void findMinuteData(Model model, Sheet sheetForMinute, ArrayList<Model> models) {
        int j = 3;
        double high;
        double low;
        Date date = new Date(model.getDateTime().getTime() + 3600000 - 60000);
        while (j < sheetForMinute.getLastRowNum()) {
            try {
                Date dateTimeOfMinuteData = sheetForMinute.getRow(j).getCell(3).getDateCellValue();
                if (dateTimeOfMinuteData.before(date) & dateTimeOfMinuteData.after(model.getDateTime())) {
                    high = sheetForMinute.getRow(j).getCell(5).getNumericCellValue();
                    low = sheetForMinute.getRow(j).getCell(6).getNumericCellValue();
                    if (high > model.getUpperBand()) {
                        if (model.getHigh() == 0.0) {
                            model.setHigh(high);
                            model.setDateTimeForHigh(dateTimeOfMinuteData);
                        }
                    }
                    if (low < model.getLowerBand()) {
                        if (model.getLow() == 0.0) {
                            model.setLow(low);
                            model.setDateTimeForLow(dateTimeOfMinuteData);
                        }
                    }
                    if (model.getLow() != 0.0 & model.getHigh() != 0.0) {
                        models.add(model);
                        return;
                    }
                }
                j++;
            } catch (NullPointerException e) {
                if (j != sheetForMinute.getLastRowNum()) {
                    j++;
                } else {
                    break;
                }
            }
        }
    }
}
