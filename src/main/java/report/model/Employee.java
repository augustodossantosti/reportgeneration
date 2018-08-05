package report.model;

import lombok.Data;

import java.util.Date;

@Data
public class Employee {

    private final String name;
    private final String email;
    private final Date dateOfBirth;
    private final Double salary;
}
