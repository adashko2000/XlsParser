package org.insert.parser.xlsParser;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.insert.parser.exceptions.WrongConfigurationException;
import org.insert.parser.sheets.SheetColumn;
import org.insert.parser.sheets.TemplateSheet;

import java.io.FileInputStream;
import java.text.MessageFormat;
import java.util.*;
import java.util.stream.Collectors;

public class XlsParser {

    private final static String INSERT_FORMAT = "INSERT INTO {0} ({1}) VALUES({2})";

    private final static String DELIMITER = ";";

    private final static String DEFAULT_VALUE = "null";

    public String convertToInsert(final TemplateSheet templateSheet) throws Exception {
        final FileInputStream file = new FileInputStream(templateSheet.filePath);
        final HSSFWorkbook workbook = new HSSFWorkbook(file);
        final Sheet sheet = workbook.getSheetAt(0);

        this.validateColumns(templateSheet, sheet);

        final Map<String, Integer> colIndexMap = this.getCellsIndex(sheet);
        final List<SheetColumn> sheetColumns = templateSheet.columns;
        final List<String> inserList = new ArrayList<>();

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (this.containsIgnored(sheet.getRow(i), templateSheet)) {
                continue;
            }
            final List<String> valueList = new ArrayList<>();
            for (final SheetColumn sheetColumn : sheetColumns) {
                final Cell cell = sheet.getRow(i).getCell(colIndexMap.get(sheetColumn.fileColName));

                if (cell != null) {
                    cell.setCellType(CellType.STRING);

                    valueList.add(this.convertToTableValue(cell, sheetColumn));
                } else {
                    valueList.add(DEFAULT_VALUE);
                }
            }

            if (!valueList.stream().allMatch(val -> val.equals("''") || val.equals("null"))) {
                inserList.add(MessageFormat.format(
                        INSERT_FORMAT,
                        templateSheet.tableName,
                        sheetColumns.stream().map(row -> row.colName).collect(Collectors.joining(",")),
                        String.join(",", valueList)));
            }

        }
        System.out.println("Insert count " + inserList.size());
        String insertListString = String.join(DELIMITER + "\n", inserList);
        return insertListString + DELIMITER;

    }

    private boolean containsIgnored(final Row row, final TemplateSheet templateSheet) {
        boolean validation = false;
        for (int i = 0; i < row.getLastCellNum(); i++) {

            final Cell cell = row.getCell(i);
            if (cell != null) {
                cell.setCellType(CellType.STRING);
                validation = templateSheet.rowIgnoringValues.contains(cell.getStringCellValue());
                if (validation) {
                    break;
                }
            }

        }
        return validation;
    }

    private String convertToTableValue(final Cell cell, final SheetColumn sheetColumn) {
        String returnValue;
        if (sheetColumn.type.equals("number")) {
            returnValue = cell.getStringCellValue();
        } else {
            returnValue = "'" + cell.getStringCellValue() + "'";
        }
        return returnValue;
    }

    private Map<String, Integer> getCellsIndex(final Sheet sheet) {
        final Map<String, Integer> map = new HashMap<>();
        final Iterator<Cell> cellIterator = sheet.getRow(0).cellIterator();
        while (cellIterator.hasNext()) {
            final Cell cell = cellIterator.next();
            map.put(cell.getStringCellValue(), cell.getColumnIndex());
        }
        return map;
    }

    private void validateColumns(final TemplateSheet templateSheet, final Sheet sheet) throws WrongConfigurationException {
        final List<String> colNames = templateSheet.columns.stream().map(col -> col.fileColName).toList();
        final List<String> fileColNames = new ArrayList<>();
        final Iterator<Cell> cellIterator = sheet.getRow(0).cellIterator();
        boolean validationResult;
        while (cellIterator.hasNext()) {
            final Cell cell = cellIterator.next();
            fileColNames.add(cell.getStringCellValue());
        }

        validationResult = fileColNames.containsAll(colNames);
        if (!validationResult) {
            throw new WrongConfigurationException("Table configuration not compatible with file: " + templateSheet.filePath);
        }
    }
}
