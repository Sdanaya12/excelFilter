package com.ExcelFilter.automation.utils;

import java.io.IOException;

public class ExcelDataCrossover {

    public static String[][] excelComparison() throws IOException, InterruptedException {
        String[][] dataOkta = ExcelOkta.excelDateComparison();
        String[][] dataICBCLA = ExcelICBCLA.excelConvertDate_ICBCLA();
        int contRow_Okta = dataOkta.length, contColumn_Okta = dataOkta[0].length;
        int contRow_ICBCLA = dataICBCLA.length, contColumn_ICBCLA = dataICBCLA[0].length;
        String [][] dataFinish = new String[contRow_Okta][contColumn_Okta];
        return dataFinish;
    }
}
