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

public class ComplexNumbers {
	
	

	
		private WebDriver driver;
		private Properties properties;

		@BeforeMethod
		public void setUp() throws IOException {
			properties = new Properties();
			String configFilePath = "C:\\Users\\Harini T M\\eclipse-workspace\\daily\\excel\\webmath\\algebra\\wordproblems.properties";
			FileInputStream fis = new FileInputStream(configFilePath);
			properties.load(fis);

			String chromeDriverPath = properties.getProperty("webdriver.chrome.driver");
	        System.setProperty("webdriver.chrome.driver", chromeDriverPath);
			driver = new ChromeDriver();
			driver.manage().window().maximize();
		}

		@Test
		public void powerOfi() {
			try {
				properties.getProperty("webmath_url");
				String powerofi = properties.getProperty("powerofi");
				properties.getProperty("expressionbox");
				properties.getProperty("input1");
				String submit = properties.getProperty("submit");

				driver.get(properties.getProperty("webmath_url"));
				driver.findElement(By.linkText(powerofi)).click();
				driver.findElement(By.xpath("expressionbox")).sendKeys("input1");
				driver.findElement(By.xpath(submit)).click();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		@Test
		public void add() {
			try {
				properties.getProperty("algebra");
				String add = properties.getProperty("add");
				properties.getProperty("expressionbox");
				properties.getProperty("input2");
				String submit = properties.getProperty("submit");

				driver.get(properties.getProperty("webmath_url"));
				driver.findElement(By.linkText(add)).click();
				driver.findElement(By.xpath("expressionbox")).sendKeys("input2");
				driver.findElement(By.xpath(submit)).click();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		public void subtract() {
			try {
				properties.getProperty("algebra");
				String subtract = properties.getProperty("subtract");
				properties.getProperty("expressionbox");
				properties.getProperty("input3");
				String submit = properties.getProperty("submit");

				driver.get(properties.getProperty("webmath_url"));
				driver.findElement(By.linkText(subtract)).click();
				driver.findElement(By.xpath("expressionbox")).sendKeys("input3");
				driver.findElement(By.xpath(submit)).click();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		public void multiply() {
			try {
				properties.getProperty("algebra");
				String multiply = properties.getProperty("multiply");
				properties.getProperty("expressionbox");
				properties.getProperty("input4");
				String submit = properties.getProperty("submit");

				driver.get(properties.getProperty("webmath_url"));
				driver.findElement(By.linkText(multiply)).click();
				driver.findElement(By.xpath("expressionbox")).sendKeys("input4");
				driver.findElement(By.xpath(submit)).click();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

		public void crunch() {
			try {
				properties.getProperty("algebra");
				String crunch = properties.getProperty("crunch");
				properties.getProperty("expressionbox");
				properties.getProperty("input5");
				String submit = properties.getProperty("submit");

				driver.get(properties.getProperty("webmath_url"));
				driver.findElement(By.linkText(crunch)).click();
				driver.findElement(By.xpath("expressionbox")).sendKeys("input5");
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


