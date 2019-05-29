package topCountry;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

public class TopCountry {
	private ArrayList<String> countryInfo = new ArrayList<String>();
	
	public void setInfo(String info) {
		this.countryInfo.add(info);
	}
	public ArrayList<String> getInfo() {
		return countryInfo;
	}
	
	public static void main(String[] args) throws Exception {
		ArrayList<String> countryList = new ArrayList<String>();
		ArrayList<TopCountry> country = new ArrayList<TopCountry>();
		
    	System.setProperty("webdriver.chrome.driver", "chromedriver.exe");
    	
    	// Chrome options for language setup
    	ChromeOptions chromeOptions = new ChromeOptions();
	    Map<String, Object> prefs = new HashMap<String, Object>();
	    prefs.put("intl.accept_languages", "en-GB");
	    chromeOptions.setExperimentalOption("prefs", prefs);
	    WebDriver driver = new ChromeDriver(chromeOptions);

    	driver.get("https://www.internetworldstats.com/top20.htm"); 
    	driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
		
    	// Collect the name of top 20 countries
    	for(int i=3; i<23; i++) {
			WebElement tableRow = driver.findElement(By.xpath("/html/body/table[4]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr["+i+"]"));
		    WebElement tableCell = tableRow.findElement(By.xpath("/html/body/table[4]/tbody/tr/td/table/tbody/tr[1]/td/table/tbody/tr["+i+"]/td[2]"));
		    countryList.add(tableCell.getText());
		}
		
    	// Get the google.com
    	driver.get("http://www.google.com"); 
    	driver.manage().timeouts().implicitlyWait(20, TimeUnit.SECONDS);
    	
    	// Search each country
    	int i = 0;
		for (String cName : countryList) {
			country.add(new TopCountry());
			
			WebElement element = driver.findElement(By.name("q"));
			element.clear();
		    element.sendKeys(cName);
		    element.submit();
		    
		    country.get(i).setInfo(cName);
		    
		    // Info collection 
		    try {
		    	List<WebElement> elements = driver.findElements(By.xpath("//div[contains(@class, 'mod')]/div[contains(@class, 'Z1hOCe')]/div"));
		    	for (WebElement e : elements){
		    		country.get(i).setInfo(e.getText());
		    	}
		    	
		    } catch (NoSuchElementException e) {
		    	System.out.println(cName +" information Not found");
		    }
		    
		    i++;
		}
	    
	    driver.close();
	    
	    
	    // Write excel sheet
	    XSSFWorkbook workbook = new XSSFWorkbook(); 
    	XSSFSheet sheet = workbook.createSheet("Top country");
    	
    	int maxRow = 0;
    	Row row = null;
    	row = sheet.createRow(5);
    	for (i=0; i<countryList.size(); i++) {
    		row = sheet.getRow(5);
			Cell cell = row.createCell(i+1);
    	  	cell.setCellValue("Sample Data "+ (i+1));
    		
    		ArrayList<String> countryInfo = country.get(i).getInfo();
    		int rowNum = countryInfo.size();
    		
    		while (maxRow < rowNum) {
    			row = sheet.createRow(7+maxRow++);
    		}
    		
    		for (int j=0; j < countryInfo.size(); j++) {
    			row = sheet.getRow(7+j);
    			cell = row.createCell(i+1);
        	  	cell.setCellValue(countryInfo.get(j));
    		}
    	}
    	
	  	
	  	// Resize all columns to fit the content size
        for(i = 1; i <= countryList.size()+1; i++) {
            sheet.autoSizeColumn(i);
        }
        
		try { 
			FileOutputStream out = new FileOutputStream(new File("TopCountry.xlsx")); 
			workbook.write(out); 
			out.close();
			workbook.close();
			System.out.println("Top Country.xlsx written successfully on your project folder."); 
		} 
		catch (Exception e) { 
			e.printStackTrace(); 
		}
	}

}
