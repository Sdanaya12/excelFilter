package com.ExcelFilter.automation.pageobjects;

import com.ExcelFilter.automation.utils.ExcelICBCLA;
import com.ExcelFilter.automation.utils.ExcelDataCrossover;
import com.ExcelFilter.automation.utils.ExcelOkta;
import com.ExcelFilter.automation.utils.Times;
import net.serenitybdd.core.pages.PageObject;

import java.io.IOException;

public class RegisterPage extends PageObject {

    public RegisterPage() throws IOException {

    }
    public void ReadExcel_Okta() throws InterruptedException, IOException {
        ExcelOkta.excelFile_Okta();
        Times.waitFor(2000);
    }
    public void ReadExcel_ICBCLA() throws InterruptedException, IOException {
        ExcelICBCLA.excelFile_ICBCLA();
        Times.waitFor(1000);
    }
    public void ExcelFilter_Okta() throws InterruptedException, IOException {
        ExcelOkta.excelFilter_Okta();
        ExcelOkta.excelSeparation_Id();
        ExcelOkta.excelSeparation_LastLogin();
        ExcelOkta.excelDateComparison();
        Times.waitFor(2000);
    }

    public void ExcelFilter_ICBCLA() throws InterruptedException, IOException {
        ExcelICBCLA.excelFilter_ICBCLA();
        ExcelICBCLA.excelConvertDate_ICBCLA();
        Times.waitFor(2000);
    }

    public void ExcelVersus() throws InterruptedException, IOException {
        ExcelDataCrossover.excelComparison();
        Times.waitFor(2000);
    }
}
