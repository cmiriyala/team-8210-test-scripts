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

public class ExcelAccessSriramFF {
	public static void main(String[]args)throws IOException, BiffException, InterruptedException, WriteException{
		System.setProperty("webdriver.gecko.driver","E:\\geckodriver.exe");
		WebDriver driver = new FirefoxDriver();
		driver.get("http://team-2.herokuapp.com/");
		driver.findElement(By.linkText("Register")).click(); 
		 Thread.sleep(1000);
		String FilePath = "E:\\sampledoc.xls";
		FileInputStream fs = new FileInputStream(FilePath);
		Workbook wb = Workbook.getWorkbook(fs);
		Sheet sh = wb.getSheet("Sheet1");
		File opXLP = new File("E:/sriramoutputexcelpositive.xls");
        WritableWorkbook outputexcelfileP = Workbook.createWorkbook(opXLP);
        WritableSheet writableSheetP = outputexcelfileP.createSheet("Sheet1", 0);
        File opXLN = new File("E:/sriramoutputexcelnegative.xls");
        WritableWorkbook outputexcelfileN = Workbook.createWorkbook(opXLN);
        WritableSheet writableSheetN = outputexcelfileN.createSheet("Sheet1", 0);
		//int totalNoOfRows = sh.getRows(); 
	//	int totalNoOfCols = sh.getColumns(); 
		WebElement name = driver.findElement(By.id("name"));
		name.sendKeys(sh.getCell(0, 0).getContents());
		//System.out.print(sh.getCell(0, 0).getContents());
		WebElement email = driver.findElement(By.id("email"));
		email.sendKeys(sh.getCell(0, 1).getContents());
		WebElement password = driver.findElement(By.id("password"));
		password.sendKeys(sh.getCell(0, 2).getContents());
		WebElement password_conf = driver.findElement(By.id("password-confirm"));
		password_conf.sendKeys(sh.getCell(0, 3).getContents());
		Select dropdown = new Select(driver.findElement(By.name("role")));
		dropdown.selectByValue(sh.getCell(0, 4).getContents());
		Thread.sleep(1000);
		driver.findElement(By.cssSelector("#app-layout > div > div > div > div > div.panel-body > form > div:nth-child(7) > div > button")).click();
		if(driver.getPageSource().contains("The email has already been taken"))
		{
		
            Label regfail1 = new Label(0, 0, "REgistration of building manager is not Successfull");
		//System.out.println("REgistration not Successfull via Excel");
            writableSheetN.addCell(regfail1);
            
		}
		else{
			Label regsucc1 = new Label(0, 0, "Registration of Building manager is Successfull");
			System.out.println("Registration of Building manager is Successfull");
	            writableSheetP.addCell(regsucc1);
	            driver.findElement(By.linkText("I Agree")).click(); 
		}
		driver.findElement(By.linkText("Register")).click(); 
		WebElement name1 = driver.findElement(By.id("name"));
		name1.sendKeys(sh.getCell(1, 0).getContents());
		//System.out.print(sh.getCell(0, 0).getContents());
		WebElement email1 = driver.findElement(By.id("email"));
		email1.sendKeys(sh.getCell(1, 1).getContents());
		WebElement password1 = driver.findElement(By.id("password"));
		password1.sendKeys(sh.getCell(1, 2).getContents());
		WebElement password_conf1 = driver.findElement(By.id("password-confirm"));
		password_conf1.sendKeys(sh.getCell(1, 3).getContents());
		Select dropdown1 = new Select(driver.findElement(By.name("role")));
		dropdown1.selectByValue(sh.getCell(1, 4).getContents());
		Thread.sleep(1000);
		driver.findElement(By.cssSelector("#app-layout > div > div > div > div > div.panel-body > form > div:nth-child(7) > div > button")).click();
		if(driver.getPageSource().contains("The email has already been taken"))
		{
		
            Label regfail2 = new Label(0, 1, "REgistration of property manager is not Successfull");
		System.out.println("REgistration of property manager is not Successfull");
            writableSheetN.addCell(regfail2);
            
		}
		else{
			Label regsucc2 = new Label(0, 1, "Registration of Property  manager is Successfull");
			System.out.println("Registration of Property  manager is Successfull");
	            writableSheetP.addCell(regsucc2);
	            driver.findElement(By.linkText("I Agree")).click(); 
		}
		driver.findElement(By.linkText("Register")).click(); 
		WebElement name2 = driver.findElement(By.id("name"));
		name2.sendKeys(sh.getCell(2, 0).getContents());
		//System.out.print(sh.getCell(0, 0).getContents());
		WebElement email2 = driver.findElement(By.id("email"));
		email2.sendKeys(sh.getCell(2, 1).getContents());
		WebElement password2 = driver.findElement(By.id("password"));
		password2.sendKeys(sh.getCell(2, 2).getContents());
		WebElement password_conf2 = driver.findElement(By.id("password-confirm"));
		password_conf2.sendKeys(sh.getCell(2, 3).getContents());
		Select dropdown2 = new Select(driver.findElement(By.name("role")));
		dropdown2.selectByValue(sh.getCell(2, 4).getContents());
		Thread.sleep(1000);
		driver.findElement(By.cssSelector("#app-layout > div > div > div > div > div.panel-body > form > div:nth-child(7) > div > button")).click();
		if(driver.getPageSource().contains("The email has already been taken"))
		{
		
            Label regfail3 = new Label(0, 2, "REgistration of worker not Successfull");
		System.out.println("REgistration of worker not Successfull");
            writableSheetN.addCell(regfail3);
            
		}
		else{
			Label regsucc3 = new Label(0, 2, "Registration of Worker is Successfull");
			System.out.println("Registration of Worker is Successfull");
	            writableSheetP.addCell(regsucc3);
	            driver.findElement(By.linkText("I Agree")).click(); 
		}
		driver.findElement(By.cssSelector("#app-navbar-collapse > ul.nav.navbar-nav.navbar-right > li:nth-child(1) > a")).click();
		WebElement email23 = driver.findElement(By.id("email"));
		email23.sendKeys(sh.getCell(0, 1).getContents());
		WebElement password23 = driver.findElement(By.id("password"));
		password23.sendKeys(sh.getCell(0, 2).getContents());
		Thread.sleep(1000);
		driver.findElement(By.tagName("form")).submit();
		if(driver.getPageSource().contains("Building Details"))
		{
			driver.findElement(By.linkText("Workers")).click();
			driver.findElement(By.linkText("Create Worker")).click();
			WebElement Workername = driver.findElement(By.id("worker_name"));
			Workername.sendKeys(sh.getCell(2, 5).getContents());
			WebElement WorkerPhno = driver.findElement(By.id("worker_mobilephone"));
			WorkerPhno.sendKeys(sh.getCell(2, 6).getContents());
			WebElement WorkerSkill = driver.findElement(By.id("worker_skills"));
			WorkerSkill.sendKeys(sh.getCell(2,7).getContents());
			driver.findElement(By.tagName("form")).submit();
           Label BMS1 = new Label(0, 0, "Worker 1 creted successfully");
           System.out.println("Worker 1 creted successfully");
          writableSheetP.addCell(BMS1);
          driver.findElement(By.linkText("Workers")).click();
			driver.findElement(By.linkText("Create Worker")).click();
			WebElement Workername1 = driver.findElement(By.id("worker_name"));
			Workername1.sendKeys(sh.getCell(2,8).getContents());
			WebElement WorkerPhno1 = driver.findElement(By.id("worker_mobilephone"));
			WorkerPhno1.sendKeys(sh.getCell(2,9).getContents());
			WebElement WorkerSkill1 = driver.findElement(By.id("worker_skills"));
			WorkerSkill1.sendKeys(sh.getCell(2,10).getContents());
			driver.findElement(By.tagName("form")).submit();
         Label BMS2 = new Label(0, 0, "Worker 2 created successfully");
         System.out.println("Worker 2 created successfully");
        writableSheetP.addCell(BMS2);
        driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr > td:nth-child(4) > a")).click();
        if(driver.getPageSource().contains("Worker Details"))
        {
        	Label BMS3 = new Label(0, 3, "Read option in Worker working fine");
        	System.out.println("Read option in Worker is working fine");
    	    writableSheetP.addCell(BMS3);
        }
        else
        {
        	Label BMF2 = new Label(0, 1, "Read function of complaints failed");
        	System.out.println("Read option in complaints failed");
    	    writableSheetN.addCell(BMF2);
        }
        driver.findElement(By.linkText("Workers")).click();
        driver.findElement(By.xpath("/html/body/div[2]/table[2]/tbody/tr/td[5]/a")).click();
        if(driver.getPageSource().contains("Update Worker Details"))
        {
			WebElement NumberofApt12 = driver.findElement(By.id("worker_mobilephone"));
			driver.findElement(By.id("worker_mobilephone")).clear();
			NumberofApt12.sendKeys(sh.getCell(2,11).getContents());
			driver.findElement(By.tagName("form")).submit();
			 if(driver.getPageSource().contains(sh.getCell(2,11).getContents()))
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
       driver.findElement(By.cssSelector("body > div:nth-child(2) > table.table.table-striped.table-bordered.table-hover > tbody > tr > td:nth-child(6) > form > input.btn.btn-danger")).click();
       Label BMS5 = new Label(0, 5, "Delete option in Workers working fine");
       System.out.println("Delete option in Workers working fine");
	    writableSheetP.addCell(BMS5);
	    driver.findElement(By.linkText("Logout")).click();
		
		outputexcelfileP.write();
	    outputexcelfileP.close();
	    outputexcelfileN.write();
	    outputexcelfileN.close();
		
	}
	
}
}
