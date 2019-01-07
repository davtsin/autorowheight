package ru.davtsin.jxls;

import org.jxls.builder.xls.XlsCommentAreaBuilder;
import org.jxls.common.Context;
import org.jxls.util.JxlsHelper;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;

/**
 * Object collection output demo
 *
 * @author Leonid Vysochyn
 */
public class ObjectCollectionDemo {
    private static Logger logger = LoggerFactory.getLogger(ObjectCollectionDemo.class);

    public static void main(String[] args) throws ParseException, IOException {
        logger.info("Running Object Collection demo");
        List<Employee> employees = generateSampleEmployeeData();
        try (InputStream is = ObjectCollectionDemo.class.getResourceAsStream("object_collection_template_merged.xls")) {
            try (OutputStream os = new FileOutputStream("target/object_collection_output.xls")) {

                XlsCommentAreaBuilder.addCommandMapping("autoRowHeight", AutoRowHeightCommand.class);

                Context context = new Context();
                context.putVar("employees", employees);
                JxlsHelper.getInstance().processTemplate(is, os, context);
            }
        }
    }

    public static List<Employee> generateSampleEmployeeData() throws ParseException {
        List<Employee> employees = new ArrayList<Employee>();
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMM-dd", Locale.US);
        employees.add(new Employee("Mr. Bolton, making a visit to Israel, told reporters that American forces would remain in Syria until the last remnants of the Islamic State were defeated and Turkey provided guarantees that it would not strike Kurdish forces allied with the United States. He and other top White House advisers have led a behind-the-scenes effort to slow Mr. Trump’s order and reassure allies, including Israel.", dateFormat.parse("1970-Jul-10"), 1500, 0.15));
        employees.add(new Employee("WASHINGTON — President Trump’s national security adviser, John R. Bolton, rolled back on Sunday Mr. Trump’s decision to rapidly withdraw from Syria, laying out conditions for a pullout that could leave American forces there for months or even years.", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
        employees.add(new Employee("Mr. Bolton’s comments inserted into Mr. Trump’s strategy something the president had omitted when he announced on Dec. 19 that the United States would depart within 30 days: any conditions that must be met before the pullout.", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
        employees.add(new Employee("While Mr. Bolton said on Sunday that he expected American forces to eventually leave northeastern Syria, where most of the 2,000 troops in the country are based for the mission against the Islamic State, he began to lay out an argument for keeping some troops at a garrison in the southeast that is used to monitor the flow of Iranian arms and soldiers. In September, three months before Mr. Trump’s announcement, Mr. Bolton had declared that the United States would remain in Syria as long as Iranians were on the ground there.", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
        employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
        return employees;
    }
}
