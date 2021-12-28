package com.ExcelFilter.automation.steps;

import com.ExcelFilter.automation.pageobjects.RegisterPage;
import com.ExcelFilter.automation.utils.Times;
import net.thucydides.core.annotations.Step;

import java.io.IOException;

public class RegisterPageSteps {

    RegisterPage registerPage = new RegisterPage();
    public RegisterPageSteps() throws IOException {
    }

    @Step
    public void openExcel_Okta() throws InterruptedException, IOException {
        registerPage.ReadExcel_Okta();
        Times.waitFor(1000);
        registerPage.ExcelFilter_Okta();
        Times.waitFor(1000);
    }
    @Step
    public void openExcel_ICBCLA() throws InterruptedException, IOException {
        registerPage.ReadExcel_ICBCLA();
        Times.waitFor(1000);
        registerPage.ExcelFilter_ICBCLA();
        Times.waitFor(1000);
    }
    @Step
    public void saveExcel() throws InterruptedException, IOException {
        registerPage.ExcelVersus();
        Times.waitFor(1000);
    }
}
