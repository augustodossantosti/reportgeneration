package report.model;

import java.time.Duration;
import java.time.Instant;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class EmployeeRepository {

    public List<Employee> findAll() {
        return createEmployees();
    }

    private List<Employee> createEmployees() {
        List<Employee> employees = new ArrayList<>();
        for (int i = 1; i <= 20; i++) {
            employees.add(new Employee("Employee" + i, "Email" + i, Date.from(Instant.now().plus(Duration.ofDays(i))), 200d * i));
        }
        return employees;
    }
}
