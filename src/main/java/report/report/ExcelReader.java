package report.report;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import report.model.Employee;

import java.io.File;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class ExcelReader {

    public List<Employee> readExcel(final ReportProperties properties) throws IOException, InvalidFormatException {
        final Workbook workbook = WorkbookFactory.create(new File(properties.getFilepath()));
        final Sheet sheet = workbook.getSheet(properties.getSheetName());
        final List<Employee> employees = new ArrayList<>();
        sheet.forEach(row -> {
            try {
                if (row.getRowNum() != 0) {
                    employees.add(createEmployee(row, properties.getDateFormat()));
                }
            } catch (ParseException e) {
                e.printStackTrace();
            }
        });
        return Collections.unmodifiableList(employees);
    }

    private Employee createEmployee(final Row row, final String dateFormat) throws ParseException {
        final DataFormatter formatter = new DataFormatter();
        final SimpleDateFormat format = new SimpleDateFormat(dateFormat);
        return new Employee(
                formatter.formatCellValue(row.getCell(0)),
                formatter.formatCellValue(row.getCell(1)),
                format.parse(formatter.formatCellValue(row.getCell(2))),
                Double.valueOf(formatter.formatCellValue(row.getCell(3)))
        );
    }
}
