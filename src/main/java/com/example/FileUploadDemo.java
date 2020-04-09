package com.example;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.text.NumberFormat;

public class FileUploadDemo {

    public static void main(String[] args) throws Exception {
        FileInputStream fileInputStream = new FileInputStream("src/main/resources/【到家-KA】永旺门店数据4.5.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        XSSFSheet sheet = workbook.getSheetAt(0);

        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            XSSFRow row1 = sheet.getRow(i);
            for (int j = 0; j < row1.getLastCellNum(); j++) {
                XSSFCell cell = row1.getCell(j);
                String cellValue = getCellValue(cell);
                System.out.print(cellValue + " ");
            }
            System.out.println();
        }
    }

    private static String getCellValue(XSSFCell cell) {
        String cellValue;
        if(cell.getCellType().equals(CellType.NUMERIC)) {
            NumberFormat numberFormat = NumberFormat.getInstance();
            cellValue = numberFormat.format(cell.getNumericCellValue());
            if (cellValue.indexOf(",") >= 0) {
                cellValue = cellValue.replace(",", "");
            }
        }else {
            cellValue = cell.getStringCellValue();
        }
        return cellValue;
    }
}
