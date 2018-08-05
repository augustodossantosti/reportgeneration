package report;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import report.model.Employee;
import report.model.EmployeeRepository;
import report.report.ExcelReader;
import report.report.ExcelWriter;
import report.report.ReportProperties;

import java.io.IOException;
import java.util.List;

public class AppInitializer {

    public static void main(String ...args) throws IOException, InvalidFormatException {

        // Prepara os dados a serem escritos no relatorio
        final EmployeeRepository repository = new EmployeeRepository();
        final List<Employee> employees = repository.findAll();

        // Gera o relatorio de Employee e escreve o arquivo excel no disco
        final ReportProperties properties = new ReportProperties();
        new ExcelWriter().writeReport(employees, properties);

        // Le o arquivo gerado e constroi uma lista de Employee
        final List<Employee> employeesFromFile = new ExcelReader().readExcel(properties);
        System.out.println(employeesFromFile);
    }
}
