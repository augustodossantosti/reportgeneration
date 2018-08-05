package report.report;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import report.model.Employee;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.IntStream;

public class ExcelWriter {

    private Workbook workbook;

    public void writeReport(final List<Employee> data, final ReportProperties properties) throws IOException {
        workbook = new XSSFWorkbook();
        final Sheet sheet = workbook.createSheet(properties.getSheetName());
        writeHeader(sheet, properties.getColumns());
        writeData(sheet, properties, data);
        writeFileOnDisk(properties.getFilepath());
    }

    private void writeHeader(final Sheet sheet, final String[] headerColumns) {
        final int HEADER_ROW_INDEX = 0;
        final Row headerRow = sheet.createRow(HEADER_ROW_INDEX);
        for (int i = 0; i < headerColumns.length; i++) {
            final Cell cell = headerRow.createCell(i);
            cell.setCellValue(headerColumns[i]);
            cell.setCellStyle(getHeaderCellStyle());
        }
    }

    private CellStyle getHeaderCellStyle() {
        final Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 14);
        headerFont.setColor(IndexedColors.RED.getIndex());
        final CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        return headerCellStyle;
    }

    private CellStyle getDateCellStyle(final String dateFormat) {
        final CreationHelper helper = workbook.getCreationHelper();
        final CellStyle dateCellStyle = workbook.createCellStyle();
        dateCellStyle.setDataFormat(helper.createDataFormat().getFormat(dateFormat));
        return dateCellStyle;
    }

    private void writeData(final Sheet sheet, final ReportProperties properties, final List<Employee> data) {
        int DATA_START_ROW = 1;
        for (final Employee employee : data) {
            final Row row = sheet.createRow(DATA_START_ROW++);
            row.createCell(0).setCellValue(employee.getName());
            row.createCell(1).setCellValue(employee.getEmail());
            final Cell dataOfBirthCell = row.createCell(2);
            dataOfBirthCell.setCellValue(employee.getDateOfBirth());
            dataOfBirthCell.setCellStyle(getDateCellStyle(properties.getDateFormat()));
            row.createCell(3).setCellValue(employee.getSalary());
        }
        resizeAllColumns(sheet, properties.getColumns().length);
    }

    private void resizeAllColumns(final Sheet sheet, final int columns) {
        IntStream.range(0, columns).forEach(sheet::autoSizeColumn);
    }

    private void writeFileOnDisk(final String filepath) throws IOException {
        final FileOutputStream fileOut = new FileOutputStream(filepath);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }

}
