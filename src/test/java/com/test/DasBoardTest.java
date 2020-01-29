package com.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.PropertyConfigurator;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.testng.Assert;
import org.testng.annotations.BeforeClass;
import org.testng.annotations.BeforeSuite;
import org.testng.annotations.DataProvider;

import org.testng.annotations.Test;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

public class DasBoardTest {
	WebDriver driver = null;
	WebElement we = null;
	List<WebElement> listLinksWebEle = null;
	ArrayList<String> listMenuWeb = null;
	ArrayList<String> listMenuExcel = null;
	public static String PathOfLogFile = "C:\\Users\\Shree\\Desktop\\Alleclipse\\kiran\\DashBoardTesting\\log4j.properties";
	public static org.apache.log4j.Logger log;

	@BeforeSuite
	public void openBrowser() {
		System.setProperty("webdriver.chrome.driver",
				"C:\\Users\\Shree\\Desktop\\Alleclipse\\kiran\\DashBoardTesting\\src\\test\\java\\com\\drivers\\chromedriver.exe");
		driver = new ChromeDriver();
		driver.get("file:///E:/Offline%20Website/Offline%20Website/pages/examples/dashboard.html");

		log = org.apache.log4j.Logger.getLogger(DasBoardTest.class);
		PropertyConfigurator.configure(PathOfLogFile);
	}

	@BeforeClass
	public void loadMenuForTesting() {
		listMenuWeb = new ArrayList<>();
		listMenuExcel = new ArrayList<>();
		listLinksWebEle = driver.findElements(By.xpath("//section//ul//li//a"));

		// listLinksWebEle = we.findElements(By.tagName("li"));
		for (WebElement webElement : listLinksWebEle) {
			listMenuWeb.add(webElement.getText());
		}
		listMenuExcel = getExcelList("C:\\Users\\Shree\\Desktop\\MainLinks.xls", "Mainmenu");

	}

	@Test(dataProvider = "empLogin", priority = 1)
	public void testLinkSpellings(String linkNameExcel) {
		Assert.assertTrue(listMenuWeb.contains(linkNameExcel));
		log.info(" Spelling of Webtableas links text" + listMenuWeb + "Is Matching With Exceltable  " + listMenuExcel);
	}

	@Test(priority = 2)
	public void testLinksCount() {
		Assert.assertTrue(listMenuWeb.size() == listMenuExcel.size());
		log.info("list of MenuWeb is matching with list of MenuExcel");
	}

	@Test(priority = 3)
	public void testLinksSeq() {
		Assert.assertEquals(listMenuWeb, listMenuExcel);
		if (listMenuWeb.size() == listMenuExcel.size()) {
			for (int i = 1; i < listMenuWeb.size(); i++) {
				boolean flag = true;
				if (!listMenuWeb.get(i).equals(listMenuExcel.get(i))) {
					flag = false;
					break;
				}
				Assert.assertTrue(flag);
				log.info("Links Sequence is correct");

			}
		} else {
			Assert.assertTrue(false);
			log.info("Links Sequence is not correct");
		}
	}

	@DataProvider(name = "empLogin")
	public Object[][] loginData() {
		Object[][] arrayObject = getExcelData("C:\\Users\\Shree\\Desktop\\MainLinks.xls", "Mainmenu");
		return arrayObject;
	}

	public ArrayList<String> getExcelList(String fileName, String sheetName) {
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet("Mainmenu");
			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();

			for (int i = 1; i < totalNoOfRows; i++) {
				for (int j = 0; j < totalNoOfCols; j++) {
					listMenuExcel.add(sh.getCell(j, i).getContents());
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return listMenuExcel;
	}

	public String[][] getExcelData(String fileName, String sheetName) {
		String[][] arrayExcelData = null;
		try {
			FileInputStream fs = new FileInputStream(fileName);
			Workbook wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet("Mainmenu");

			int totalNoOfCols = sh.getColumns();
			int totalNoOfRows = sh.getRows();

			arrayExcelData = new String[totalNoOfRows - 1][totalNoOfCols];

			for (int i = 1; i < totalNoOfRows; i++) {

				for (int j = 0; j < totalNoOfCols; j++) {
					arrayExcelData[i - 1][j] = sh.getCell(j, i).getContents();
				}

			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			e.printStackTrace();
		} catch (BiffException e) {
			e.printStackTrace();
		}
		return arrayExcelData;
	}
}
