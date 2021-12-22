package com.demoWinium.automation.stepdefinitions;

import com.demoWinium.automation.steps.RegisterPageSteps;
import io.cucumber.java.en.Given;
import io.cucumber.java.en.Then;
import io.cucumber.java.en.When;
import net.thucydides.core.annotations.Steps;

import java.io.IOException;

public class DemoWiniumStepDefinitions {
    @Steps
    RegisterPageSteps registerPageSteps;

    @Given("A user of the work team selects the Excel document")
    public void a_user_of_the_work_team_selects_the_excel_document() throws IOException, InterruptedException {
        registerPageSteps.openExcel();
    }

    @When("The system performs the reading of the document")
    public void the_system_performs_the_reading_of_the_document() throws IOException, InterruptedException {
        registerPageSteps.filterExcel();
    }

    @Then("You should see the document properly filtered")
    public void you_should_see_the_document_properly_filtered() throws IOException, InterruptedException {
        registerPageSteps.saveExcel();
    }
}
