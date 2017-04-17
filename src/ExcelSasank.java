import org.openqa.selenium.*;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelSasank {
	public static void main(String[]args)throws IOException, BiffException, InterruptedException, WriteException{
		System.setProperty("webdriver.chrome.driver", "E:\\chromedriver.exe");
		WebDriver driver = new ChromeDriver();
		driver.get("http://team-2.herokuapp.com");
		driver.findElement(By.cssSelector("#app-navbar-collapse > ul.nav.navbar-nav.navbar-right > li:nth-child(1) > a")).click(); 
		 Thread.sleep(1000);
		String FilePath = "E:\\sampledocsasank.xls";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet("Sheet1");
		File opXLP = new File("E:/sasankoutputexcelpositive.xls");
        WritableWorkbook outputexcelfileP = Workbook.createWorkbook(opXLP);
        WritableSheet writableSheetP = outputexcelfileP.createSheet("Sheet1", 0);
        File opXLN = new File("E:/sasankoutputexcelnegative.xls");
        WritableWorkbook outputexcelfileN = Workbook.createWorkbook(opXLN);
        WritableSheet writableSheetN = outputexcelfileN.createSheet("Sheet1", 0);
		WebElement email = driver.findElement(By.id("email"));
		email.sendKeys(sh.getCell(0, 1).getContents());
		WebElement password = driver.findElement(By.id("password"));
		password.sendKeys(sh.getCell(0, 2).getContents());
		Thread.sleep(1000);
		driver.findElement(By.tagName("form")).submit();
		
		
		if(driver.getPageSource().contains("Building Details"))
		{
			Label BMS = new Label(0, 0, "Login as building manager is successfully");
	        writableSheetP.addCell(BMS);
	        System.out.println("Login as a building manager Successfull");
			driver.findElement(By.linkText("Complaints")).click();
			driver.findElement(By.linkText("Create Complaint")).click();
			Select dropdown = new Select(driver.findElement(By.name("building_id")));
			dropdown.selectByVisibleText(sh.getCell(0,5).getContents());
			WebElement Customername = driver.findElement(By.id("customer_name"));
			Customername.sendKeys(sh.getCell(0, 6).getContents());
			WebElement CustomerEmail = driver.findElement(By.id("customer_email"));
			CustomerEmail.sendKeys(sh.getCell(0, 7).getContents());
			WebElement CustMobNo = driver.findElement(By.id("customer_mobilephone"));
			CustMobNo.sendKeys(sh.getCell(0,8).getContents());
			WebElement CustAPTNO = driver.findElement(By.id("customer_aptphone"));
			CustAPTNO.sendKeys(sh.getCell(0,9).getContents());
			WebElement CustComplaint = driver.findElement(By.id("customer_complaint"));
			CustComplaint.sendKeys(sh.getCell(0,10).getContents());
			WebElement CustAddress = driver.findElement(By.id("customer_address"));
			CustAddress.sendKeys(sh.getCell(0,11).getContents());
			driver.findElement(By.tagName("form")).submit();
           Label BMS1 = new Label(0, 0, "Complaint 1 creted successfully");
           System.out.println("Complaint 1 created successfully");
          writableSheetP.addCell(BMS1);
			driver.findElement(By.linkText("Complaints")).click();
			driver.findElement(By.linkText("Create Complaint")).click();
			Select dropdown1 = new Select(driver.findElement(By.name("building_id")));
			dropdown1.selectByVisibleText(sh.getCell(0,12).getContents());
			WebElement Customername1 = driver.findElement(By.id("customer_name"));
			Customername1.sendKeys(sh.getCell(0, 13).getContents());
			WebElement CustomerEmail1 = driver.findElement(By.id("customer_email"));
			CustomerEmail1.sendKeys(sh.getCell(0, 14).getContents());
			WebElement CustMobNo1 = driver.findElement(By.id("customer_mobilephone"));
			CustMobNo1.sendKeys(sh.getCell(0,15).getContents());
			WebElement CustAPTNO1 = driver.findElement(By.id("customer_aptphone"));
			CustAPTNO1.sendKeys(sh.getCell(0,16).getContents());
			WebElement CustComplaint1 = driver.findElement(By.id("customer_complaint"));
			CustComplaint1.sendKeys(sh.getCell(0,17).getContents());
			WebElement CustAddress1 = driver.findElement(By.id("customer_address"));
			CustAddress1.sendKeys(sh.getCell(0,18).getContents());
			driver.findElement(By.tagName("form")).submit();
         Label BMS2 = new Label(0, 0, "Complaint 2 created successfully");
         System.out.println("Complaint 2 created successfully");
        writableSheetP.addCell(BMS2);
        driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(8) > a")).click();
        if(driver.getPageSource().contains("Complaint Details"))
        {
        	Label BMS3 = new Label(0, 3, "Read option in complaints working fine");
        	System.out.println("Read option in complaints is working fine");
    	    writableSheetP.addCell(BMS3);
        }
        else
        {
        	Label BMF2 = new Label(0, 1, "Read function of complaints failed");
        	System.out.println("Read option in complaints failed");
    	    writableSheetN.addCell(BMF2);
        }
        driver.findElement(By.linkText("Complaints")).click();
        driver.findElement(By.xpath("/html/body/div[2]/table[2]/tbody/tr[1]/td[9]/a")).click();
        if(driver.getPageSource().contains("Update Complaint"))
        {
			WebElement NumberofApt12 = driver.findElement(By.id("customer_mobilephone"));
			driver.findElement(By.id("customer_mobilephone")).clear();
			NumberofApt12.sendKeys(sh.getCell(0,19).getContents());
			driver.findElement(By.tagName("form")).submit();
			 if(driver.getPageSource().contains(sh.getCell(0,19).getContents()))
		        {
        	Label BMS4 = new Label(0, 4, "Update option in Complaints working fine");
        	System.out.println("Update option in Complaints working fine");
    	    writableSheetP.addCell(BMS4);
		        }
        }
        else
        {
        	Label BMF4 = new Label(0, 2, "update function of Complaints failed");
        	System.out.println("Update option in Complaints not working fine");
    	    writableSheetN.addCell(BMF4);
        }
       driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(10) > form > input.btn.btn-danger")).click();
       Label BMS5 = new Label(0, 5, "Delete option in Complaints working fine");
       System.out.println("Delete option in Complaints working fine");
	    writableSheetP.addCell(BMS5);
	    driver.findElement(By.linkText("Logout")).click();
       Label BMS6 = new Label(0, 5, "Logout as Property Manager is Successfull");
       System.out.println("Logout as Property Manager Successfull");
	    writableSheetP.addCell(BMS6); 
		}
		else{
		Label BMF1 = new Label(0, 0, "Creation of Complaints failed");
	    writableSheetN.addCell(BMF1); 
	}
		outputexcelfileP.write();
	    outputexcelfileP.close();
	    outputexcelfileN.write();
	    outputexcelfileN.close();
		
	}
}
