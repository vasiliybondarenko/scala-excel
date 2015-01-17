package poi.excel.report.examples;

import poi.excel.report.api.*;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

/**
 * Created with IntelliJ IDEA.
 * Author: shredinger
 * Date: 1/16/15
 * Time: 5:16 PM
 * Project: scala-excel
 */
public class ExcelReporterSamples {
    private DataField acTypeCol= new DataField("aircraftType", DataType.String);
    private DataField saCol = new DataField("serviceArea", DataType.String);
    private DataField contractMins = new DataField("minutesContracted", DataType.Number);
    private DataField usedMins = new DataField("minutesUsed", DataType.Number);
    private DataField revenue = new DataField("revenue", DataType.Number);
    private DataField reportDate = new DataField("reportDate", DataType.Date);
    private DataField reportMoney = new DataField("reportMoney", DataType.Money);

    private ReportField acTypeRepField = new ReportField("aircraftType", "AC Type");
    private ReportField saRF = new ReportField("serviceArea", "Service Area");
    private ReportField contractMinsRF = new ReportField("minutesContracted", "Contract Minutes");
    private ReportField usedMinsRF = new ReportField("minutesUsed", "Used Minutes");
    private ReportField revenueRF = new ReportField("revenue", "Revenue", "0.00");
    private ReportField reportDateRF = new ReportField("reportDate", "Report Date");
    private ReportField reportMoneyRF = new ReportField("reportMoney", "$");

    List<DataField> dataFields = Arrays.asList(new DataField[]{
            acTypeCol, saCol, contractMins, usedMins, revenue, reportDate, reportMoney
    });

    List<List<String>> data = new ArrayList(){{
        add(Arrays.asList("AT1", "SA2", "2000", "1000", "1000000", "2015/01/15", "2000"));
        add(Arrays.asList("AT2", "SA2", "2500", "1200", "2000000", "2015/01/14", "123000"));
        add(Arrays.asList("AT3", "SA2", "3400", "1700", "1800000", "2015/01/12", "123000"));
        add(Arrays.asList("AT1", "SA2", "2100", "1100", "2500000", "2015/02/01", "123000"));
        add(Arrays.asList("AT2", "SA1", "2400", "1300", "2300000", "2015/01/28", "123000"));
        add(Arrays.asList("AT3", "SA2", "3200", "1500", "2100000", "2015/01/21", "123000"));
    }};

    public void simpleReport(){
        ReportInfo repInfo = new ReportInfo(dataFields,
                Arrays.asList(acTypeRepField, saRF, contractMinsRF, revenueRF, reportDateRF, reportMoneyRF),
                new ArrayList<ReportField>()
        );
        new ExcelReporter().run(repInfo, data);
    }

    public void groupReport(){
        ReportInfo repInfo = new ReportInfo(dataFields,
                Arrays.asList(acTypeRepField, saRF, contractMinsRF, revenueRF, reportDateRF, reportMoneyRF),
                Arrays.asList(acTypeRepField, saRF)
        );
        new ExcelReporter().run(repInfo, data);
    }

    public static void main(String[] args) {
        new ExcelReporterSamples().simpleReport();
    }

}
