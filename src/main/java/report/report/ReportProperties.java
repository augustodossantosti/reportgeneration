package report.report;

import lombok.Getter;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;

@Getter
public class ReportProperties {

    public ReportProperties() throws IOException {
        final String reportPropertiesPath = Thread.currentThread().getContextClassLoader().getResource("report.properties").getPath();
        final Properties reportProperties = new Properties();
        reportProperties.load(new FileInputStream(reportPropertiesPath));
        sheetName = reportProperties.getProperty("sheetName");
        filepath = reportProperties.getProperty("filepath");
        columns = reportProperties.getProperty("columns").split(",");
        dateFormat = reportProperties.getProperty("dateFormat");
    }

    private String[] columns;
    private String sheetName;
    private String filepath;
    private String dateFormat;
}
