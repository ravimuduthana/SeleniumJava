import java.text.ParseException;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

public class PrintTableData 
{
	    public static void main(String[] args) throws ParseException
	    {
	    	System.setProperty("webdriver.chrome.driver", "C:\\Users\\RAVI\\eclipse-workspace\\chromedriver.exe");	
	    	WebDriver driver=new ChromeDriver();
	    	driver.get("http://demo.guru99.com/test/web-table-element.php");         
	        //No.of Columns
	        List  col = driver.findElements(By.xpath(".//*[@id=\"leftcontainer\"]/table/thead/tr/th"));
	        System.out.println("No of cols are : " +col.size()); 
	        //No.of rows 
	        List  rows = driver.findElements(By.xpath(".//*[@id='leftcontainer']/table/tbody/tr/td[1]")); 
	        System.out.println("No of rows are : " + rows.size());
	        driver.close();
	    }
}
