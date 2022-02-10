package com.outlook.ilyakastsenevich.excelgenerator;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import lombok.experimental.UtilityClass;
import org.apache.commons.lang3.reflect.FieldUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

@UtilityClass
public class ExcelGenerator {

  //generate 1 sheet xlsx
  public byte[] generateExcel(List<?> dtoDataList, Class<?> dtoClass) {
    XSSFWorkbook workbook = new XSSFWorkbook();
    XSSFSheet sheet = workbook.createSheet();

    createHeaders(sheet, dtoClass);

    XSSFRow row;

    int rowNumber = 1;

    for (Object object : dtoDataList) {
      row = sheet.createRow(rowNumber);

      int cellNumber = 0;

      Cell numberCell = row.createCell(cellNumber++);
      numberCell.setCellValue(rowNumber++);

      for (Field field : object.getClass().getDeclaredFields()) {
        Cell cell = row.createCell(cellNumber++);
        String value = null;

        try {
          value = String.valueOf(FieldUtils.readDeclaredField(object, field.getName(), true));
        } catch (IllegalAccessException e) {
          throw new RuntimeException(e);
        }

        cell.setCellValue(value);
      }

    }

    if (sheet.getPhysicalNumberOfRows() > 0) {
      Row sheetRow = sheet.getRow(sheet.getFirstRowNum());
      Iterator<Cell> cellIterator = sheetRow.cellIterator();
      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        Integer columnIndex = cell.getColumnIndex();
        sheet.autoSizeColumn(columnIndex);
      }
    }

    byte[] doc;

    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      workbook.write(outputStream);
      doc = outputStream.toByteArray();
    } catch (IOException e) {
      throw new RuntimeException("Failed to create document", e);
    }

    return doc;
  }

  //generate multiple sheet xlsx
  //key = sheet name
  //value = dto class to list with dto data pair
  public byte[] generateExcel(Map<String, Pair<Class<?>, List<?>>> sheetNameToDtoClassToDtoDataPairMap) {
    XSSFWorkbook workbook = new XSSFWorkbook();

    for(Map.Entry<String, Pair<Class<?>, List<?>>> entry: sheetNameToDtoClassToDtoDataPairMap.entrySet()) {
      createSheet(workbook, entry.getKey(), entry.getValue().getRight(), entry.getValue().getLeft());
    }

    byte[] doc;

    try (ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {
      workbook.write(outputStream);
      doc = outputStream.toByteArray();
    } catch (IOException e) {
      throw new RuntimeException("Failed to create document", e);
    }

    return doc;
  }

  private void createSheet(XSSFWorkbook workbook, String sheetName, List<?> dtoDataList, Class<?> dtoClass) {
    XSSFSheet sheet = workbook.createSheet(sheetName);

    createHeaders(sheet, dtoClass);

    XSSFRow row;

    int rowNumber = 1;

    for (Object object : dtoDataList) {
      row = sheet.createRow(rowNumber);

      int cellNumber = 0;

      Cell numberCell = row.createCell(cellNumber++);
      numberCell.setCellValue(rowNumber++);

      for (Field field : object.getClass().getDeclaredFields()) {
        Cell cell = row.createCell(cellNumber++);
        String value = null;

        try {
          value = String.valueOf(FieldUtils.readDeclaredField(object, field.getName(), true));
        } catch (IllegalAccessException e) {
          throw new RuntimeException(e);
        }

        cell.setCellValue(value);
      }

    }

    if (sheet.getPhysicalNumberOfRows() > 0) {
      Row sheetRow = sheet.getRow(sheet.getFirstRowNum());
      Iterator<Cell> cellIterator = sheetRow.cellIterator();
      while (cellIterator.hasNext()) {
        Cell cell = cellIterator.next();
        Integer columnIndex = cell.getColumnIndex();
        sheet.autoSizeColumn(columnIndex);
      }
    }
  }

  private void createHeaders(XSSFSheet spreadsheet, Class<?> clazz) {
    XSSFRow row = spreadsheet.createRow(0);

    int cellNumber = 0;
    Cell numberCell = row.createCell(cellNumber++);
    numberCell.setCellValue("№");

    Field[] fields = clazz.getDeclaredFields();

    for (Field field : fields) {
      Cell cell = row.createCell(cellNumber++);
      cell.setCellValue(field.getName());
    }
  }
}
