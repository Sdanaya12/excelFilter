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
    public void openExcel() throws InterruptedException, IOException {
        registerPage.ReadExcel();
        Times.waitFor(1000);
    }
    @Step
    public void filterExcel() throws InterruptedException, IOException {
        registerPage.ExcelFilter();
        Times.waitFor(1000);
    }
    @Step
    public void saveExcel() throws InterruptedException, IOException {

        Times.waitFor(1000);
    }
}
