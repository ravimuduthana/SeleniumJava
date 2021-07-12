import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
public class GetToolTipText 
{
	
	public static void main(String[] args)
	{
		System.setProperty("webdriver.chrome.driver", "C:\\Users\\RAVI\\eclipse-workspace\\chromedriver.exe");	
		WebDriver driver=new ChromeDriver();
		driver.get("https://demoqa.com/tool-tips/");
		driver.manage().window().maximize();
		String strtext = driver.findElement(By.id("toolTipButton")).getText();
		System.out.println(strtext);
		if(strtext.equals("Hover me to see")) 
		{							
            System.out.println("Test Case Passed");					
        }		
        driver.close();
	}
	
}
