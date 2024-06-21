package testRunner;

import org.junit.runner.RunWith;

import io.cucumber.junit.Cucumber;
import io.cucumber.junit.CucumberOptions;

@RunWith(Cucumber.class)
@CucumberOptions(
				features= {".//Features"},
//				features={".//Features/TS_001.feature",".//Features/TS_002.feature"},
//				features= {".//Features/TS_004_DD.feature"},
//		features= {".//Features/DDTest.feature"},
//							features= {"@target/rerun.txt"},
					glue="stepDefinitions",
					plugin= {"pretty", "html:reports/myreport.html", 
							  "rerun:target/rerun.txt",
							  "com.aventstack.extentreports.cucumber.adapter.ExtentCucumberAdapter:"
							},
					dryRun=false,    // checks mapping between scenario steps and step definition methods
					monochrome=true,    // to avoid junk characters in output
					publish=true
					//,tags="@smoke"
		)
public class TestRunner {

		}
