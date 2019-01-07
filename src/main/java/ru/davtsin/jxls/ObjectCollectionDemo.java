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
        employees.add(new Employee("Elsaaaaaaaaaaaaaaaaaaa aaaaaaaaaaaaaaaaaaaa aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa aaaaaaaaaaa", dateFormat.parse("1970-Jul-10"), 1500, 0.15));
        employees.add(new Employee("Oleg gggg gggggggg kkkkkkkkkkkkkkk kkkkkkkkkk", dateFormat.parse("1973-Apr-30"), 2300, 0.25));
        employees.add(new Employee("Neil llllllllll lllllllll ggggggggggggggggggg lllllllllllllll", dateFormat.parse("1975-Oct-05"), 2500, 0.00));
        employees.add(new Employee("Maria aaaa aaaa aaaaaaaaaa", dateFormat.parse("1978-Jan-07"), 1700, 0.15));
        employees.add(new Employee("John", dateFormat.parse("1969-May-30"), 2800, 0.20));
        return employees;
    }
}
