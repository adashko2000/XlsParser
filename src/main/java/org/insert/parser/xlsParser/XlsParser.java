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

public class XlsParser {

    private final static String INSERT_FORMAT = "INSERT INTO {0} ({1}) VALUES({2})";

    private final static String DELIMITER = ";";

    private final static String DEFAULT_VALUE = "null";

    private final static String ID = "ID";

    public String convertToInsert(final TemplateSheet templateSheet) throws Exception {
        final FileInputStream file = new FileInputStream(templateSheet.filePath);
        final HSSFWorkbook workbook = new HSSFWorkbook(file);
        final Sheet sheet = workbook.getSheetAt(0);

        this.validateColumns(templateSheet, sheet);

        final Map<String, Integer> colIndexMap = this.getCellsIndex(sheet);
        final List<SheetColumn> sheetColumns = templateSheet.columns;
        final List<String> inserList = new ArrayList<>();

        long tableSequence = 1;

        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            if (this.omitCells(sheet.getRow(i), templateSheet, sheet.getRow(0))) {
                continue;
            }
            final List<String> valueList = new ArrayList<>();
            for (final SheetColumn sheetColumn : sheetColumns) {
                final Cell cell = sheet.getRow(i).getCell(colIndexMap.get(sheetColumn.fileColName));

                if (cell != null) {
                    cell.setCellType(CellType.STRING);
                    if (cell.getStringCellValue() != "") {
                        valueList.add(this.convertToTableValue(cell, sheetColumn));
                    } else {
                        valueList.add(DEFAULT_VALUE);
                    }
                } else {
                    valueList.add(DEFAULT_VALUE);
                }
            }

            if (!valueList.stream().allMatch(val -> val.equals("''") || val.equals("null"))) {
                final List<String> insertColNames = new ArrayList<>();
                insertColNames.add(ID);
                insertColNames.addAll(sheetColumns.stream().map(row -> row.colName).toList());

                final List<String> insertValueList = new ArrayList<>();
                insertValueList.add(String.valueOf(tableSequence));
                insertValueList.addAll(valueList);

                inserList.add(MessageFormat.format(
                        INSERT_FORMAT,
                        templateSheet.tableName,
                        String.join(",", insertColNames),
                        String.join(",", insertValueList)));

                tableSequence++;
            }

        }
        System.out.println("Insert count " + inserList.size());
        String insertListString = String.join(DELIMITER + "\n", inserList);
        return insertListString + DELIMITER;

    }

    private boolean omitCells(final Row row, final TemplateSheet templateSheet, final Row headerRow) {
        boolean validation = false;
        for (int i = 0; i < row.getLastCellNum(); i++) {

            final Cell cell = row.getCell(i);
            final Cell headerCell = headerRow.getCell(i);

            if (cell != null && headerCell != null) {
                cell.setCellType(CellType.STRING);
                validation = this.containsIgnored(cell, templateSheet, headerCell.getStringCellValue())
                        || !this.requiredFieldNotEmpty(cell, templateSheet, headerCell.getStringCellValue());
                if (validation) {
                    break;
                }
            }

        }
        return validation;
    }

    private boolean containsIgnored(final Cell cell, final TemplateSheet templateSheet, final String headerName) {
        final Optional<SheetColumn> sheetColumn = templateSheet.columns.stream()
                .filter(col -> col.fileColName.equals(headerName))
                .findFirst();

        return sheetColumn.isPresent() && templateSheet.rowIgnoringValues.contains(cell.getStringCellValue());
    }

    private boolean requiredFieldNotEmpty(final Cell cell, final TemplateSheet templateSheet, final String headerName) {
        final Optional<SheetColumn> sheetColumn = templateSheet.columns.stream()
                .filter(col -> col.fileColName.equals(headerName))
                .findFirst();
        if (!sheetColumn.isPresent()) {
            return true;
        }
        return sheetColumn.get().required ? cell.getStringCellValue() != "" : true;
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
