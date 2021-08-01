package scripts;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileOutputStream;
import java.io.FileReader;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.jayway.jsonpath.JsonPath;

import net.minidev.json.JSONArray;

public class jsonpoc 
{
	public static void main(String[] args)
	{
        try
        {     
        	BufferedReader reader=new BufferedReader(new FileReader(".\\JsonData.json"));
            
        	
        	
        	String line=null;
        	StringBuilder builder=new StringBuilder();
        	while((line=reader.readLine())!=null)
        	{
        		//System.out.println(line);
        		builder.append(line);
        	}
        	String jsoncontent=builder.toString();
        	//System.out.println(jsoncontent);
        	
        	//Assign element jsonpath to json arrays
        	JSONArray arraycourse=JsonPath.read(jsoncontent, "$.courses..course");
        	JSONArray arrayduration=JsonPath.read(jsoncontent, "$.courses..duration");
        	JSONArray arraycourseid=JsonPath.read(jsoncontent, "$.courses..courseId");
        	JSONArray arrayusersname=JsonPath.read(jsoncontent, "$.users..name");
        	JSONArray arrayuserseid=JsonPath.read(jsoncontent, "$.users..eid");
        	JSONArray arrayusershobby=JsonPath.read(jsoncontent, "$.users..hobby");
        	JSONArray arrayuserscourseid=JsonPath.read(jsoncontent, "$.users..courseId");
        	JSONArray arrayusersaddress1=JsonPath.read(jsoncontent, "$.users..address.address1");
        	JSONArray arrayuserscity=JsonPath.read(jsoncontent, "$.users..address.city");
        	JSONArray arrayusersstate=JsonPath.read(jsoncontent, "$.users..address.state");
        	JSONArray arrayuserszipcode=JsonPath.read(jsoncontent, "$.users..address.zipCode");
        	JSONArray arrayuserscountry=JsonPath.read(jsoncontent, "$.users..address.country");
        	
        	//Assign arraylengths to variables
        	int courselength=arraycourse.size();
        	int durationlength=arrayduration.size();
        	int courseidlength=arraycourseid.size();
        	int arrayusersnamelength=arrayusersname.size();
        	int arrayuserseidlength=arrayuserseid.size();
        	int arrayusershobbylength=arrayusershobby.size();
        	int arrayuserscourseidlength=arrayuserscourseid.size();
        	int arrayusersaddress1length=arrayusersaddress1.size();
        	int arrayuserscitylength=arrayuserscity.size();
        	int arrayusersstatelength=arrayusersstate.size();
        	int arrayuserszipcodelength=arrayuserszipcode.size();
        	int arrayuserscountrylength=arrayuserscountry.size();
        	
        	//create an object for the Jsonexcel
        	File file =    new File(".\\JsonExcel.xlsx");
          	FileOutputStream outputStream = new FileOutputStream(file);
          	XSSFWorkbook  jasonexcel = new XSSFWorkbook ();
          	XSSFSheet mySheet = jasonexcel.createSheet("Json");
          	
          	int i;
          	
          	//Get course and write into excel
          	//validate course and write into excel
          	
        		for(i=0;i<courselength;i++)
        		{
        			XSSFRow myrow=mySheet.createRow(i);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellcourse=myrow.createCell(1);
        			Cell cellcoursestatus=myrow.createCell(2);
        			Cell cellcoursevalidation=myrow.createCell(3);
        			cellcourse.setCellValue(arraycourse.get(i).toString());
        			cellvalname.setCellValue("course");
        			String strcourse=arraycourse.get(i).toString();
        			strcourse=strcourse.replaceAll("\\s+","");
        			if(strcourse.matches("[a-zA-Z]+"))
        			{
        				if(strcourse.length()<=7)
        				{
        					cellcoursestatus.setCellValue("Valid");
        					cellcoursevalidation.setCellValue("It is a Alphabetic and length less than or equal to 7");
        				}
        				else
        				{
        					
        					cellcoursestatus.setCellValue("In Valid");
        					cellcoursevalidation.setCellValue("It is a Alphabetic and length greater than 7");
        				}
        			}
        			else
        			{
        				
        				cellcoursestatus.setCellValue("In Valid");
        				cellcoursevalidation.setCellValue("It is not Alphabetic");
        				
        			}
                    //System.out.println(arraycourse.get(i));
                   
                    //cell.setCellValue(arraycourse.get(i));
        		}
        		int activerow=i;
        		
        		//Get duration and write into excel
        		//validate duration and write into excel
        		
        		for(int j=0;j<durationlength;j++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellduration=myrow.createCell(1);
        			Cell celldurationstatus=myrow.createCell(2);
        			Cell celldurationvalidation=myrow.createCell(3);
        			cellduration.setCellValue(arrayduration.get(j).toString());
        			
        			cellvalname.setCellValue("duration");
        			String strduration=arrayduration.get(j).toString();
        			double d;
        			d=0.0;
        			try
        			{
        				d=Double.parseDouble(strduration);
        				if((d % 1)==0)
        				{
	        				
	            				celldurationstatus.setCellValue("In Valid");
	            				celldurationvalidation.setCellValue("Value is Not Double");
	            			
        				}
        				else
        				{
        					if(d>3 && d<1)
	            			{
	        					celldurationstatus.setCellValue("In Valid");
	        					celldurationvalidation.setCellValue("Value is not Double or not in the range 1 to 3");
	            			}
	            			else
	            			{
	            				celldurationstatus.setCellValue("Valid");
	            				celldurationvalidation.setCellValue("Value is Double and in the range 1 to 3");
	            			}
        					
        				}
        			}
        			catch(Exception e)
        			{
        				celldurationstatus.setCellValue("In Valid");
        				celldurationvalidation.setCellValue("Value is Not Double");
        			}
        			
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get courseid and write into excel
        		//validate courseid and write into excel
        		
        		for(int k=0;k<courseidlength;k++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellcourseid=myrow.createCell(1);
        			Cell cellcourseidstatus=myrow.createCell(2);
        			Cell cellcourseidvalidation=myrow.createCell(3);
        			cellcourseid.setCellValue(arraycourseid.get(k).toString());
        			
        			cellvalname.setCellValue("courseid");
        			
        			String strcourseid=arraycourseid.get(k).toString();
        			if(StringUtils.isNumeric(strcourseid))
        			{
        				if(strcourseid.length()<=1)
        				{
        					cellcourseidstatus.setCellValue("Valid");
        					cellcourseidvalidation.setCellValue("It is Numeric and length less than or equal to 1");
        				}
        				else
        				{
        					cellcourseidstatus.setCellValue("In Valid");
        					cellcourseidvalidation.setCellValue("It is not Numeric and length greater than 1");
        				}
        			}
        			else
        			{
        				
        					
        				cellcourseidstatus.setCellValue("In Valid");
        				cellcourseidvalidation.setCellValue("It is not Numeric");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get name and write into excel
        		//validate name and write into excel
        		
        		for(int name=0;name<arrayusersnamelength;name++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellname=myrow.createCell(1);
        			Cell cellnamestatus=myrow.createCell(2);
        			Cell cellnamevalidation=myrow.createCell(3);
        			cellname.setCellValue(arrayusersname.get(name).toString());
        			
        			cellvalname.setCellValue("name");
        			
        			String strname=arrayusersname.get(name).toString();
        
        			strname=strname.replaceAll("\\s+","");
        			if(strname.matches("[a-zA-Z]+"))
        			{
        				
        				
        				if(strname.length()<30)
        				{
        					cellnamestatus.setCellValue("Valid");
        					cellnamevalidation.setCellValue("It is a Alphabetic and length less than 30");
        				}
        				else
        				{
        					cellnamestatus.setCellValue("In Valid");
        					cellnamevalidation.setCellValue("It is a Alphabetic and length greater than 30");
        				}
        			}
        			else
        			{
        				
        				
        					cellnamestatus.setCellValue("In Valid");
        					cellnamevalidation.setCellValue("It is not  Alphabetic");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get eid and write into excel
        		//validate eid and write into excel
        		
        		for(int eid=0;eid<arrayuserseidlength;eid++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell celleid=myrow.createCell(1);
        			Cell celleidstatus=myrow.createCell(2);
        			Cell celleidvalidation=myrow.createCell(3);
        			celleid.setCellValue(arrayuserseid.get(eid).toString());
        			
        			cellvalname.setCellValue("eid");
        			
        			String streid=arrayuserseid.get(eid).toString();
        			if(StringUtils.isNumeric(streid))
        			{
        				if(streid.length()==5)
        				{
        					celleidstatus.setCellValue("Valid");
        					celleidvalidation.setCellValue("Value is Numeric and length is fixed to 5");
        				}
        				else
        				{
        					celleidstatus.setCellValue("In Valid");
        					celleidvalidation.setCellValue("Value is Numeric and length is not fixed to 5");
        				}
        			}
        			else
        			{
        					celleidstatus.setCellValue("In Valid");
        					celleidvalidation.setCellValue("Value is not Numeric");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get hobby and write into excel
        		//validate hobby and write into excel
        		
        		for(int hobby=0;hobby<arrayusershobbylength;hobby++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellhobby=myrow.createCell(1);
        			Cell cellhobbystatus=myrow.createCell(2);
        			Cell cellhobbyvalidation=myrow.createCell(3);
        			cellhobby.setCellValue(arrayusershobby.get(hobby).toString());
        			
        			cellvalname.setCellValue("hobby");
        			
        			String strhobby=arrayusershobby.get(hobby).toString();
        
        			strhobby=strhobby.replaceAll("\\s+","");
        			if(strhobby.matches("[a-zA-Z]+"))
        			{
        				
        				
        				if(strhobby.length()<15)
        				{
        					cellhobbystatus.setCellValue("Valid");
        					cellhobbyvalidation.setCellValue("It is a Alphabetic and length less than 15");
        				}
        				else
        				{
        					cellhobbystatus.setCellValue("In Valid");
        					cellhobbyvalidation.setCellValue("It is a Alphabetic and length greater than 15");
        				}
        			}
        			else
        			{
        				
        				
        					cellhobbystatus.setCellValue("In Valid");
        					cellhobbyvalidation.setCellValue("It is not  Alphabetic");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get users.courseid and write into excel
        		//validate users.courseid and write into excel
        		
        		for(int userscourseid=0;userscourseid<arrayuserscourseidlength;userscourseid++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			
        			Cell cellvalname=myrow.createCell(0);
        			Cell celluserscourseid=myrow.createCell(1);
        			Cell celluserscourseidstatus=myrow.createCell(2);
        			Cell celluserscourseidvalidation=myrow.createCell(3);
        			celluserscourseid.setCellValue(arrayuserscourseid.get(userscourseid).toString());
        			
        			cellvalname.setCellValue("users course id");
        			
        			String struserscourseid=arraycourseid.get(userscourseid).toString();
        			if(StringUtils.isNumeric(struserscourseid))
        			{
        				if(struserscourseid.length()<=1)
        				{
        					celluserscourseidstatus.setCellValue("Valid");
        					celluserscourseidvalidation.setCellValue("It is Numeric and length less than or equal to 1");
        				}
        				else
        				{
        					celluserscourseidstatus.setCellValue("In Valid");
        					celluserscourseidvalidation.setCellValue("It is not Numeric and length greater than 1");
        				}
        			}
        			else
        			{
        				
        					
        				celluserscourseidstatus.setCellValue("In Valid");
        				celluserscourseidvalidation.setCellValue("It is not Numeric");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get address1 and write into excel
        		//validate address1 and write into excel
        		
        		for(int address1=0;address1<arrayusersaddress1length;address1++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			
        			Cell cellvalname=myrow.createCell(0);
        			Cell celladdress1=myrow.createCell(1);
        			Cell celladdress1status=myrow.createCell(2);
        			Cell celladdress1validation=myrow.createCell(3);
        			celladdress1.setCellValue(arrayusersaddress1.get(address1).toString());
        			
        			cellvalname.setCellValue("address1");
        			
        			String straddress1=arrayusersaddress1.get(address1).toString();
        
        			straddress1=straddress1.replaceAll("\\s+","");
        			if(straddress1.matches("[a-zA-Z0-9]+"))
        			{
        				
        				
        				if(straddress1.length()<=30)
        				{
        					celladdress1status.setCellValue("Valid");
        					celladdress1validation.setCellValue("It is a Alphanumeric and length less than 30");
        				}
        				else
        				{
        					celladdress1status.setCellValue("In Valid");
        					celladdress1validation.setCellValue("It is a Alphanumeric and length greater than 30");
        				}
        			}
        			else
        			{
        				
        					celladdress1status.setCellValue("In Valid");
        					celladdress1validation.setCellValue("It is not  Alphanumeric");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		
        		//Get city and write into excel
        		//validate city and write into excel
        		
        		for(int city=0;city<arrayuserscitylength;city++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellcity=myrow.createCell(1);
        			Cell cellcitystatus=myrow.createCell(2);
        			Cell cellcityvalidation=myrow.createCell(3);
        			cellcity.setCellValue(arrayuserscity.get(city).toString());
        			
        			cellvalname.setCellValue("city");
        			
        			String strcity=arrayuserscity.get(city).toString();
        
        			strcity=strcity.replaceAll("\\s+","");
        			if(strcity.matches("[a-zA-Z]+"))
        			{
        				
        				
        				if(strcity.length()<20)
        				{
        					cellcitystatus.setCellValue("Valid");
        					cellcityvalidation.setCellValue("It is a Alphabetic and length less than 20");
        				}
        				else
        				{
        					cellcitystatus.setCellValue("In Valid");
        					cellcityvalidation.setCellValue("It is a Alphabetic and length greater than 20");
        				}
        			}
        			else
        			{
        				
        				
        					cellcitystatus.setCellValue("In Valid");
        					cellcityvalidation.setCellValue("It is not  Alphabetic");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		//Get state and write into excel
        		//validate state and write into excel
        		
        		for(int state=0;state<arrayusersstatelength;state++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellstate=myrow.createCell(1);
        			Cell cellstatestatus=myrow.createCell(2);
        			Cell cellstatevalidation=myrow.createCell(3);
        			cellstate.setCellValue(arrayusersstate.get(state).toString());
        			
        			cellvalname.setCellValue("state");
        			
        			String strstate=arrayusersstate.get(state).toString();
        
        			strstate=strstate.replaceAll("\\s+","");
        			if(strstate.matches("[a-zA-Z]+"))
        			{
        				
        				
        				if(strstate.length()<20)
        				{
        					cellstatestatus.setCellValue("Valid");
        					cellstatevalidation.setCellValue("It is a Alphabetic and length less than 20");
        				}
        				else
        				{
        					cellstatestatus.setCellValue("In Valid");
        					cellstatevalidation.setCellValue("It is a Alphabetic and length greater than 20");
        				}
        			}
        			else
        			{
        				
        				
        					cellstatestatus.setCellValue("In Valid");
        					cellstatevalidation.setCellValue("It is not  Alphabetic");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		
        		//Get zipcode and write into excel
        		//validate zipcode and write into excel
        		
        		for(int zipcode=0;zipcode<arrayuserszipcodelength;zipcode++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellzipcode=myrow.createCell(1);
        			Cell cellzipcodestatus=myrow.createCell(2);
        			Cell cellzipcodevalidation=myrow.createCell(3);
        			cellzipcode.setCellValue(arrayuserszipcode.get(zipcode).toString());
        			
        			cellvalname.setCellValue("zipcode");
        			
        			String strzipcode=arrayuserszipcode.get(zipcode).toString();
        
        			
        			if(StringUtils.isNumeric(strzipcode))
        			{
        				
        				
        				if(strzipcode.length()==6)
        				{
        					cellzipcodestatus.setCellValue("Valid");
        					cellzipcodevalidation.setCellValue("It is a Numeric and length equal to 6");
        				}
        				else
        				{
        					cellzipcodestatus.setCellValue("In Valid");
        					cellzipcodevalidation.setCellValue("It is a Numeric and length not equal to 6");
        				}
        			}
        			else
        			{
        				
        					cellzipcodestatus.setCellValue("In Valid");
        				
        					cellzipcodevalidation.setCellValue("It is not  Numeric");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		
        		//Get country and write into excel
        		//validate country and write into excel
        		
        		for(int country=0;country<arrayuserscountrylength;country++)
        		{
        			XSSFRow myrow=mySheet.createRow(activerow);
        			Cell cellvalname=myrow.createCell(0);
        			Cell cellcountry=myrow.createCell(1);
        			Cell cellcountrystatus=myrow.createCell(2);
        			Cell cellcountryvalidation=myrow.createCell(3);
        			cellcountry.setCellValue(arrayuserscountry.get(country).toString());
        			
        			cellvalname.setCellValue("country");
        			
        			String strcountry=arrayuserscountry.get(country).toString();
        
        			strcountry=strcountry.replaceAll("\\s+","");
        			if(strcountry.matches("[a-zA-Z]+"))
        			{
        				
        				
        				if(strcountry.length()<=20)
        				{
        					cellcountrystatus.setCellValue("Valid");
        					cellcountryvalidation.setCellValue("It is a Alphabetic and length less than 20");
        				}
        				else
        				{
        					cellcountrystatus.setCellValue("In Valid");
        					cellcountryvalidation.setCellValue("It is a Alphabetic and length greater than 20");
        				}
        			}
        			else
        			{
        				
        					cellcountrystatus.setCellValue("In Valid");
        				
        					cellcountryvalidation.setCellValue("It is not  Alphabetic");
        				
        			}
        			
        			
        			
        			activerow=activerow+1;
        		}
        		
        		
        		//write all the data into excel and close all the objects
        		 jasonexcel.write(outputStream);
                 outputStream.close();
                 jasonexcel.close();
        	
           }
	
        catch (Exception e) 
        {
			System.out.println("error message is"+ e.getMessage());
		}
}
}
