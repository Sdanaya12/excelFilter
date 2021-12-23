package com.ExcelFilter.automation.pageobjects;

import com.demoWinium.automation.models.DataInjection;
import com.ExcelFilter.automation.utils.Excel;
import com.ExcelFilter.automation.utils.Times;
import net.serenitybdd.core.pages.PageObject;

import java.io.IOException;

public class RegisterPage extends PageObject {
    Excel dataInjection = new Excel();

    public RegisterPage() throws IOException {

    }

    public void ReadExcel() throws InterruptedException, IOException {
        Excel.excelFile();
        Times.waitFor(2000);
    }

    public void ExcelFilter() throws InterruptedException, IOException {
        Excel.excelFilter();
        Times.waitFor(2000);
    }
}
