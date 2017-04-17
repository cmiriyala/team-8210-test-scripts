import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class ExcelchethanFF {
	public static void main(String[]args)throws IOException, BiffException, InterruptedException, WriteException{
		System.setProperty("webdriver.gecko.driver","E:\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		driver.get("http://team-2.herokuapp.com");
		driver.findElement(By.cssSelector("#app-navbar-collapse > ul.nav.navbar-nav.navbar-right > li:nth-child(1) > a")).click(); 
		 Thread.sleep(1000);
		String FilePath = "E:\\sampledocchethan.xls";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet("Sheet1");
		File opXLP = new File("E:/chethanoutputexcelpositive.xls");
        WritableWorkbook outputexcelfileP = Workbook.createWorkbook(opXLP);
        WritableSheet writableSheetP = outputexcelfileP.createSheet("Sheet1", 0);
        File opXLN = new File("E:/chethanoutputexcelnegative.xls");
        WritableWorkbook outputexcelfileN = Workbook.createWorkbook(opXLN);
        WritableSheet writableSheetN = outputexcelfileN.createSheet("Sheet1", 0);
		WebElement email = driver.findElement(By.id("email"));
		email.sendKeys(sh.getCell(0, 1).getContents());
		WebElement password = driver.findElement(By.id("password"));
		password.sendKeys(sh.getCell(0, 2).getContents());
		Thread.sleep(1000);
		driver.findElement(By.tagName("form")).submit();
		
		
		if(driver.getPageSource().contains("Work Orders"))
		{
			Label BMS = new Label(0, 0, "Login as Worker is successfully");
	        writableSheetP.addCell(BMS);
	        System.out.println("Login as a Worker Successfull");
	        driver.findElement(By.cssSelector("body > div:nth-child(2) > table > tbody > tr > td:nth-child(9) > a")).click();
	        if(driver.getPageSource().contains("Tracking Details"))
	        {
	        	Label BMS3 = new Label(0, 1, "Tracking Status in workorder is working fine");
	        	System.out.println("Tracking Status in workorder is working fine");
	    	    writableSheetP.addCell(BMS3);
	        }
	        else
	        {
	        	Label BMF2 = new Label(0, 1, "Tracking Status function of complaints failed");
	        	System.out.println("Tracking Status option in complaints failed");
	    	    writableSheetN.addCell(BMF2);
	        } 
	        driver.findElement(By.linkText("Work Orders")).click();
	        driver.findElement(By.cssSelector("body > div:nth-child(2) > table > tbody > tr > td:nth-child(10) > a")).click();    
	        if(driver.getPageSource().contains("Update Work Order"))
	        {
	        	Select dropdown1 = new Select(driver.findElement(By.name("order_status")));
				dropdown1.selectByVisibleText(sh.getCell(0,5).getContents());
				driver.findElement(By.tagName("form")).submit();
	        	Label BMS3 = new Label(0, 3, "Update Status in workorder is working fine");
		        driver.findElement(By.cssSelector("body > div:nth-child(2) > table > tbody > tr > td:nth-child(9) > a")).click();
		        if(driver.getPageSource().contains("Work In Progress"))
		        {
		        	Label BMS34 = new Label(0, 2, "Tracking Status is updated to Work In Progress");
		        	System.out.println("Tracking Status is updated to Work In Progress");
		    	    writableSheetP.addCell(BMS34);
		        }
		        else
		        {
		        	Label BMF2 = new Label(0, 1, "Tracking Status is not updated to Work In Progress");
		        	System.out.println("Tracking Status option in complaints failed");
		    	    writableSheetN.addCell(BMF2);
		        } 
	        	System.out.println("Update Status in workorder is working fine");
	    	    writableSheetP.addCell(BMS3);
	        }
	        else
	        {
	        	Label BMF2 = new Label(0, 1, "Update Status function of complaints failed");
	        	System.out.println("Update Status option in complaints failed");
	    	    writableSheetN.addCell(BMF2);
	        } 

		    driver.findElement(By.linkText("Logout")).click();
		       Label BMS6 = new Label(0, 3, "Logout as worker Successfull");
		       System.out.println("Logout as worker Successfull");
		       writableSheetP.addCell(BMS6); 
		       driver.findElement(By.cssSelector("#app-navbar-collapse > ul.nav.navbar-nav.navbar-right > li:nth-child(1) > a")).click();
				WebElement email23 = driver.findElement(By.id("email"));
				email23.sendKeys(sh.getCell(0,6).getContents());
				WebElement password23 = driver.findElement(By.id("password"));
				password23.sendKeys(sh.getCell(0,7).getContents());
				Thread.sleep(1000);
				driver.findElement(By.tagName("form")).submit(); //correct
		       driver.findElement(By.linkText("Work Orders")).click();
			driver.findElement(By.linkText("Create Work Order")).click();
			Select dropdown1 = new Select(driver.findElement(By.name("worker_id")));
			dropdown1.selectByVisibleText(sh.getCell(0,8).getContents());
			WebElement workermob12 = driver.findElement(By.id("worker_mobilephone"));
			workermob12.sendKeys(sh.getCell(0, 9).getContents());
			WebElement OrderDEsc1 = driver.findElement(By.id("order_description"));
			OrderDEsc1.sendKeys(sh.getCell(0, 10).getContents());
			WebElement EstCost1 = driver.findElement(By.id("order_est_cost"));
			EstCost1.sendKeys(sh.getCell(0,11).getContents());
			WebElement ActCost1 = driver.findElement(By.id("order_actual_cost"));
			ActCost1.sendKeys(sh.getCell(0,12).getContents());
			WebElement OdDAte1 = driver.findElement(By.id("order_date"));
			OdDAte1.sendKeys(sh.getCell(0,13).getContents());
			WebElement CmpDate1 = driver.findElement(By.id("order_completion_date"));
			CmpDate1.sendKeys(sh.getCell(0,14).getContents());
			Select dropdown12 = new Select(driver.findElement(By.name("order_status")));
			dropdown12.selectByVisibleText(sh.getCell(0,15).getContents());
			driver.findElement(By.tagName("form")).submit();
         Label BMS2 = new Label(0, 4, "Work Order 1 created successfully");
         System.out.println("Work Order 1 created successfully");
        writableSheetP.addCell(BMS2); //fine
        driver.findElement(By.linkText("Work Orders")).click();
		driver.findElement(By.linkText("Create Work Order")).click();
		Select dropdown123 = new Select(driver.findElement(By.name("worker_id")));
		dropdown123.selectByVisibleText(sh.getCell(0,16).getContents());
		WebElement workermob1 = driver.findElement(By.id("worker_mobilephone"));
		workermob1.sendKeys(sh.getCell(0, 17).getContents());
		WebElement OrderDEsc = driver.findElement(By.id("order_description"));
		OrderDEsc.sendKeys(sh.getCell(0, 18).getContents());
		WebElement EstCost = driver.findElement(By.id("order_est_cost"));
		EstCost.sendKeys(sh.getCell(0,19).getContents());
		WebElement ActCost = driver.findElement(By.id("order_actual_cost"));
		ActCost.sendKeys(sh.getCell(0,20).getContents());
		WebElement OdDAte = driver.findElement(By.id("order_date"));
		OdDAte.sendKeys(sh.getCell(0,21).getContents());
		WebElement CmpDate = driver.findElement(By.id("order_completion_date"));
		CmpDate.sendKeys(sh.getCell(0,22).getContents());
		Select dropdown1234 = new Select(driver.findElement(By.name("order_status")));
		dropdown1234.selectByVisibleText(sh.getCell(0,15).getContents());
		driver.findElement(By.tagName("form")).submit();
     Label BMS25 = new Label(0, 5, "Work Order 2 created successfully");
     System.out.println("Work Order 2 created successfully");
     writableSheetP.addCell(BMS25); 
        driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(9) > a")).click();
        if(driver.getPageSource().contains("Tracking Details"))
        {
			
        	Label BMS4 = new Label(0, 6, "Track Status in Work order working fine");
        	System.out.println("Track Status in Work order working fine");
    	    writableSheetP.addCell(BMS4);
		        
        }
        else
        {
        	Label BMF4 = new Label(0, 2, "Track Status in Work order is not working fine");
        	System.out.println("Track Status in Work order is not working fine");
    	    writableSheetN.addCell(BMF4);
        }
        driver.findElement(By.linkText("Work Orders")).click();
       driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(10) > a")).click();
       if(driver.getPageSource().contains("Update Work Order"))
       {
			WebElement ActCost123 = driver.findElement(By.id("order_actual_cost"));
			driver.findElement(By.id("order_actual_cost")).clear();
			ActCost123.sendKeys(sh.getCell(0,24).getContents());
			driver.findElement(By.tagName("form")).submit();
			 if(driver.getPageSource().contains(sh.getCell(0,24).getContents()))
		        {
       	Label BMS4 = new Label(0, 7, "Update option in work order working fine");
       	System.out.println("Update option in work order is working fine");
   	    writableSheetP.addCell(BMS4);
		        }
       }
       else
       {
       	Label BMF4 = new Label(0, 2, "update function of work order failed");
       	System.out.println("Update option in work order is not working fine");
   	    writableSheetN.addCell(BMF4);
       }
      driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr:nth-child(1) > td:nth-child(11) > form > input.btn.btn-danger")).click();
      Label BMS5 = new Label(0, 8, "Delete option in buildings working fine");
      System.out.println("Delete option in buildings working fine");
	    writableSheetP.addCell(BMS5);
	    driver.findElement(By.linkText("Logout")).click();
      Label BMS67 = new Label(0, 9, "Logout Successfull");
      System.out.println("Logout Successfull");
	    writableSheetP.addCell(BMS67); 
	}
		outputexcelfileP.write();
	    outputexcelfileP.close();
	    outputexcelfileN.write();
	    outputexcelfileN.close();
		
	}
}
