package Assignmentprjct;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Properties;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.annotations.AfterMethod;
import org.testng.annotations.BeforeMethod;
import org.testng.annotations.Test;

public class Factorial {
	private WebDriver driver;
	private Properties properties;

	@BeforeMethod
	public void setUp() throws IOException {
		properties = new Properties();
		String configFilePath = "C:\\Users\\Harini T M\\eclipse-workspace\\daily\\excel\\Data.xlsxfactoringpolynomials.properties";
		FileInputStream fis = new FileInputStream(configFilePath);
		properties.load(fis);

		String chromeDriverPath = properties.getProperty("webdriver.chrome.driver");
        System.setProperty("webdriver.chrome.driver", chromeDriverPath);
		driver = new ChromeDriver();
		driver.manage().window().maximize();
	}

	@Test(priority = 1)
	public void usingGCF() {
		try {
			properties.getProperty("webmath_url");
			String usingGCF = properties.getProperty("usingGCF");
			String expressionbox = properties.getProperty("expressionbox");
			String input1 = properties.getProperty("input1");
			String submit = properties.getProperty("submit");

			driver.get(properties.getProperty("webmath_url"));
			driver.findElement(By.linkText(usingGCF)).click();
			driver.findElement(By.xpath(expressionbox)).sendKeys(input1);
			driver.findElement(By.xpath(submit)).click();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test(priority = 2)
	public void diffOftwoSquares() {
		try {
			properties.getProperty("webmath_url");
			String algebra = properties.getProperty("algebra");
			String diffoftwosquares = properties.getProperty("diffoftwosquares");
			String expressionbox = properties.getProperty("expressionbox");
			String input2 = properties.getProperty("input2");
			String submit = properties.getProperty("submit");

			driver.get(properties.getProperty("webmath_url"));
			driver.findElement(By.linkText(algebra)).click();
			driver.findElement(By.linkText(diffoftwosquares)).click();
            driver.findElement(By.xpath(expressionbox)).sendKeys(input2);
			driver.findElement(By.xpath(submit)).click();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@Test(priority = 3)
	public void trinomials() {
		try {
			properties.getProperty("webmath_url");
			String algebra = properties.getProperty("algebra");
			String trinomials = properties.getProperty("trinomials");
			String expressionbox = properties.getProperty("expressionbox");
			String input3 = properties.getProperty("input3");
			String submit = properties.getProperty("submit");

			driver.get(properties.getProperty("webmath_url"));
			driver.findElement(By.linkText(algebra)).click();
			driver.findElement(By.linkText(trinomials)).click();
			driver.findElement(By.xpath(expressionbox)).sendKeys(input3);
			driver.findElement(By.xpath(submit)).click();

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@AfterMethod
	private void closeApp() {
		driver.close();
	}

}


