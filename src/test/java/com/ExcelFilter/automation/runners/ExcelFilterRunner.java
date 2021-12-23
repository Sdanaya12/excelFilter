package com.ExcelFilter.automation.runners;

import io.cucumber.junit.CucumberOptions;
import net.serenitybdd.cucumber.CucumberWithSerenity;
import org.junit.runner.RunWith;

@RunWith(CucumberWithSerenity.class)
@CucumberOptions(features = "src/test/resources/features/excelFilter.feature", glue = "com.ExcelFilter.automation.stepdefinitions", tags = "@smokeTest")

public class ExcelFilterRunner {
}
