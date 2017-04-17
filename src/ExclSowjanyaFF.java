import org.openqa.selenium.*;
import org.openqa.selenium.firefox.FirefoxDriver;
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

public class ExclSowjanyaFF {
	public static void main(String[]args)throws IOException, BiffException, InterruptedException, WriteException{
		System.setProperty("webdriver.gecko.driver","E:\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		driver.get("http://team-2.herokuapp.com");
		driver.findElement(By.cssSelector("#app-navbar-collapse > ul.nav.navbar-nav.navbar-right > li:nth-child(1) > a")).click(); 
		 Thread.sleep(1000);
		String FilePath = "E:\\sampledocsowjanya.xls";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet("Sheet1");
		File opXLSP = new File("E:/sowjanyaoutputexcelpositive.xls");
        WritableWorkbook outputexcelfilePO = Workbook.createWorkbook(opXLSP);
        WritableSheet writableSheetPO = outputexcelfilePO.createSheet("Sheet1", 0);
        File opXLSN = new File("E:/sowjanyaoutputexcelnegative.xls");
        WritableWorkbook outputexcelfileNE = Workbook.createWorkbook(opXLSN);
        WritableSheet writableSheetNE = outputexcelfileNE.createSheet("Sheet1", 0);
		WebElement email = driver.findElement(By.id("email"));
		email.sendKeys(sh.getCell(0, 1).getContents());
		WebElement password = driver.findElement(By.id("password"));
		password.sendKeys(sh.getCell(0, 2).getContents());
		Thread.sleep(1000);
		driver.findElement(By.tagName("form")).submit();
		
		if(driver.getPageSource().contains("Property Details"))
		{
			Label BMS0 = new Label(0, 0, "Login as a property manager Successfull");
		       System.out.println("Login as a property manager Successfull");
			    writableSheetPO.addCell(BMS0);
			//rowCount=driver.findElements(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr")).size();
			//System.out.println(rowCount);
			driver.findElement(By.linkText("Buildings")).click();
			driver.findElement(By.linkText("Add Building")).click();
			Select dropdown = new Select(driver.findElement(By.name("name")));
			dropdown.selectByVisibleText(sh.getCell(0,5).getContents());
			WebElement Buildingname = driver.findElement(By.id("building_name"));
			Buildingname.sendKeys(sh.getCell(0, 6).getContents());
			WebElement BuildingAddress = driver.findElement(By.id("building_address"));
			BuildingAddress.sendKeys(sh.getCell(0, 7).getContents());
			WebElement NumberofApt = driver.findElement(By.id("number_of_apartments"));
			NumberofApt.sendKeys(sh.getCell(0,8).getContents());
			driver.findElement(By.tagName("form")).submit();
           Label BMS1 = new Label(0, 1, "1st Building creted successfully");
           System.out.println("1st Building creted successfully");
          writableSheetPO.addCell(BMS1);      
          driver.findElement(By.linkText("Buildings")).click();
			driver.findElement(By.linkText("Add Building")).click();
			Select dropdown12 = new Select(driver.findElement(By.name("name")));
			dropdown12.selectByVisibleText(sh.getCell(0,9).getContents());
			WebElement Buildingname1 = driver.findElement(By.id("building_name"));
			Buildingname1.sendKeys(sh.getCell(0, 10).getContents());
			WebElement BuildingAddress1 = driver.findElement(By.id("building_address"));
			BuildingAddress1.sendKeys(sh.getCell(0, 11).getContents());
			WebElement NumberofApt1 = driver.findElement(By.id("number_of_apartments"));
			NumberofApt1.sendKeys(sh.getCell(0,12).getContents());
			driver.findElement(By.tagName("form")).submit();
         Label BMS2= new Label(0, 2, "2nd Building creted successfully");
         System.out.println("2nd Building creted successfully");
        writableSheetPO.addCell(BMS2);
        driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(4) > a")).click();
        if(driver.getPageSource().contains("Building Details"))
        {
        	Label BMS3 = new Label(0, 3, "Read option in buildings working fine");
        	System.out.println("Read option in buildings working fine");
    	    writableSheetPO.addCell(BMS3);
        }
        else
        {
        	Label BMF2 = new Label(0, 1, "Read function of Buildings failed");
        	System.out.println("Read option in buildings failed");
    	    writableSheetNE.addCell(BMF2);
        }
        driver.findElement(By.linkText("Buildings")).click();
        driver.findElement(By.xpath("/html/body/div[2]/table[2]/tbody/tr[1]/td[5]/a")).click();
        if(driver.getPageSource().contains("Update Building Details"))
        {
			WebElement NumberofApt12 = driver.findElement(By.id("number_of_apartments"));
			driver.findElement(By.id("number_of_apartments")).clear();
			NumberofApt12.sendKeys(sh.getCell(0,13).getContents());
			driver.findElement(By.tagName("form")).submit();
			 if(driver.getPageSource().contains(sh.getCell(0,13).getContents()))
		        {
        	Label BMS4 = new Label(0, 4, "Update option in buildings working fine");
        	System.out.println("Update option in buildings working fine");
    	    writableSheetPO.addCell(BMS4);
		        }
        }
        else
        {
        	Label BMF4 = new Label(0, 2, "update function of Buildings failed");
        	System.out.println("Update option in buildings not working fine");
    	    writableSheetNE.addCell(BMF4);
        }
       driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(6) > form > input.btn.btn-danger")).click();
       Label BMS5 = new Label(0, 5, "Delete option in buildings working fine");
       System.out.println("Delete option in buildings working fine");
	    writableSheetPO.addCell(BMS5);
	    //driver.findElement(By.cssSelector("body > div:nth-child(2) > table:nth-child(2) > tbody > tr > td:nth-child(2) > a > button")).click();
	   // Label BMS8 = new Label(0, 6, "Download of the report Successfull");
	     //  System.out.println("Download of the report Successfull");
		 //   writableSheetPO.addCell(BMS8); 
	    driver.findElement(By.linkText("Logout")).click();
       Label BMS6 = new Label(0, 5, "Logout Successfull");
       System.out.println("Logout Successfull");
	    writableSheetPO.addCell(BMS6); 
		}
		else{
		Label BMF1 = new Label(0, 0, "Creation of Buildings failed");
	    writableSheetNE.addCell(BMF1);
	   //         driver.findElement(By.linkText("I Agree")).click(); 
	}
		outputexcelfilePO.write();
	    outputexcelfilePO.close();
	    outputexcelfileNE.write();
	    outputexcelfileNE.close();
		
	}
}
		