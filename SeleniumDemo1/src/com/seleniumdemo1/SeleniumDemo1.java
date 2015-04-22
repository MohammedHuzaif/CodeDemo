package com.seleniumdemo1;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.Properties;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.read.biff.BiffException;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

import org.apache.commons.io.FileUtils;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
//import java.text.DateFormat;
//import java.text.SimpleDateFormat;
//import jxl.write.Label;
//import jxl.write.WritableCell;


public class SeleniumDemo1 {

	/*Objectives:
	 * - Read from Test case Spread Sheet and execute test cases which are 'Pending'
	 * - Change status of test cases as per the results after execution
	 * - Open Gmail
	 * - Check Login with different passwords
	 * - Login and perform clicks
	 * - Open Spread Sheet in Google Drive and compare with local file
	 */
		static WebDriver driver;
		static WebElement element;
		private String trueFile, sharedFile, testCaseFile, testCaseFileTemp;
		private int trueCopyRowCount, 
					sharedCopyRowCount, 
					//trueCopyColumnCount, 
					sharedCopyColumnCount;

		static Properties prop = new Properties();
		static SeleniumDemo1 SDC=new SeleniumDemo1();
		
		public static void main(String[] args) throws Exception{
			
			System.out.println("public static void main()");
		
			String 	testCaseData=null,
					description=null;
					//testCaseStatus=null;
			//String elementLocator=null;
			Boolean status=false;
			
			SDC.getProperties();
			
			SDC.setTrueFilePath(prop.getProperty("trueCopyLocalSystemPath"));
			SDC.setSharedFilePath(prop.getProperty("sharedCopyLocalSystemPath"));
			SDC.trueCopyRowCount=Integer.parseInt(prop.getProperty("trueCopyRowCount"));
			SDC.sharedCopyRowCount=Integer.parseInt(prop.getProperty("sharedCopyRowCount"));
			//SDC.trueCopyColumnCount=Integer.parseInt(prop.getProperty("trueCopyColumnCount"));
			SDC.sharedCopyColumnCount=Integer.parseInt(prop.getProperty("sharedCopyColumnCount"));
			SDC.setTestCaseFilePath(prop.getProperty("testCaseSheetPath"));
			SDC.setTestCaseFilePathTemp(prop.getProperty("testCaseSheetPathTemp"));
					
			int 	rowCount=1, 
					//columnCount=0,
					maxRowCount=Integer.parseInt(prop.getProperty("testCaseMaxRowCount"));
					//maxColumnCount=Integer.parseInt(prop.getProperty("testCaseMaxColCount"))
					
			
			//**************************************************//
			//Open Browser
			SDC.openBrowser();
			//*************************************************//
					
			while(rowCount<=maxRowCount){
				
			description=SDC.readTestCaseSheet(rowCount,7);
			System.out.println(description);
			
			
			if(description.equals("Pending") || description.equals("pending") || description.equals("PENDING") || description.equals("") || description.equals(" ") || description.equals(null)){
				testCaseData=SDC.readTestCaseSheet(rowCount,1);
				System.out.println(description+" "+rowCount+" "+testCaseData);
			}
			else{
				description=null;
				testCaseData="passed";
				System.out.println(description+" + "+rowCount);
				}
			
			switch (testCaseData){
			
			case "Open GDrive":
				System.out.println("switch-case: Open GDrive");
				status=true;
				status=SDC.openGoogleDriveLink();;
				SDC.setStatusToTestCaseSheet(status, rowCount);
				break;
			
			case "Login":
				System.out.println("switch-case: Login");
				status=true;
				status=SDC.logIn();
				SDC.setStatusToTestCaseSheet(status, rowCount);
				break;
				
			case "Perform Clicks":
				System.out.println("switch-case: Perform Clicks");
				status=false;
				//status=SDC.performClicks();
				SDC.setStatusToTestCaseSheet(status, rowCount);
				break;
				
			case "Compare Gdrive Sheet with Local Sheet":
				System.out.println("switch-case: Compare Gdrive Sheet with Local Sheet");
				status=true;
				status=SDC.compareGDriveSheetWithLocalSheet();
				SDC.setStatusToTestCaseSheet(status, rowCount);
				break;
					
			default :
				System.out.println("switch-case: default");
				break;

			}

			rowCount++;
			}
			
			//SDC.compareLocalSheets();
			//SDC.editSheet(1, 2,"cellData","GREEN");
			//SDC.callEditSheet();		
			//SDC.readTrueCopy(0, 0);
			//SDC.readSharedCopy(1, 1);
			//SDC.performKeyboardActionsOnGDriveDoc();	
			
			//**************************************************//
			//Close Browser
			SDC.closeBrowser();
			//**************************************************//
			
			System.out.println("End of Automation Testing");
		}
		
			public void getProperties() {
				
				System.out.println("getProperties()");
			
				//Properties prop = new Properties();
				InputStream propInputFile = null;	
				
				try {
					 
					//Change the prop.properties filepath as per the direstory location in your systems
					propInputFile = new FileInputStream("C:\\Users\\hassantabrez\\git\\SeleniumDemo\\SeleniumDemo1\\src\\com\\seleniumdemo1\\prop.properties");
			 
					// load a properties file
					prop.load(propInputFile);
			 
					/*
					// get the property value and print it out
					System.out.println("gmailurl : "+prop.getProperty("gmailurl"));
					System.out.println("emailid : "+prop.getProperty("emailid"));
					System.out.println("passwrd1 : "+prop.getProperty("passwrd1"));
					System.out.println("passwrd2 : "+prop.getProperty("passwrd2"));
					System.out.println("passwrd3 : "+prop.getProperty("passwrd3"));
					System.out.println("passwrd4 : "+prop.getProperty("passwrd4"));
					System.out.println("passwrd5 : "+prop.getProperty("passwrd5"));
					System.out.println("passwrd6 : "+prop.getProperty("passwrd6"));
					System.out.println("sharedCopyGDriveLink : "+prop.getProperty("sharedCopyGDriveLink"));
					System.out.println("trueCopyLocalSystemPath : "+prop.getProperty("trueCopyLocalSystemPath"));
					System.out.println("testCaseSheetPath : "+prop.getProperty("testCaseSheetPath"));
					*/
					
				} catch (IOException ex) {
					ex.printStackTrace();
				} finally {
					if (propInputFile != null) {
						try {
							propInputFile.close();
						} catch (IOException e) {
							e.printStackTrace();
						}
					}
				}
				
			}
				
			public void openBrowser() throws Exception {
				
				System.out.println("openBrowser()");
			
					//Launch the Firefox Browser
					System.out.println("Starting up the Firefox Browser");
					driver = new FirefoxDriver();
					
					System.out.println("Sleeping (2 sec)");
			        Thread.sleep(2000);
					
					String windowHandleSharedCopy = driver.getWindowHandle(); 
			        String windowTitleSharedCopy =driver.getTitle();
			        System.out.println("The Shared Document URL (windowHandleSharedCopy) : " + driver.getCurrentUrl()+" "+windowHandleSharedCopy );
			        System.out.println("The Shared Document Window (windowTitleSharedCopy) : " + windowTitleSharedCopy );

			        //Maximize the browser window
			        System.out.println("Maximizing the browser window");
			        driver.manage().window().maximize();
			        
			}
			
			public Boolean openGoogleDriveLink() throws Exception{
				Boolean mStatus=false;
				System.out.println("openGoogleDriveLink()");
			
				System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);
				
				//open Gmail.com
				System.out.println("Opening the Google Drive website");
				System.out.println(prop.getProperty("sharedCopyGDriveLink"));
				driver.get(prop.getProperty("sharedCopyGDriveLink"));
				
				System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);	
		        
		        mStatus=true;
		        System.out.println(mStatus);
		        return mStatus;
			}
			
			public Boolean openGmail() throws Exception{
				Boolean mStatus=false;
				System.out.println("openGmail()");
			
				System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);
				
				//open Gmail.com
				System.out.println("Opening the gmail.com website");
				System.out.println(prop.getProperty("gmailurl"));
				driver.get(prop.getProperty("gmailurl"));
				
				mStatus=driver.getPageSource().contains("One account. All of Google.");
				System.out.println(mStatus);
				
				System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);	
		        
		        return mStatus;
			}
			
			public Boolean logIn() throws Exception{
				Boolean mStatus=false;
				System.out.println("logIn()");
				
				System.out.println("Enter User Email");
		        
		        //Username
		        driver.findElement(By.id("Email")).sendKeys(prop.getProperty("emailid"));
		        
		        System.out.println("Enter User Password");
		        
		      //Password
		        
		        driver.findElement(By.id("Passwd")).sendKeys("xtreamitppl");
		      //Click Sign In button
		        driver.findElement(By.id("signIn")).click();
		        
		        System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);
		        
		        //Password
		        for(int k=1;k<=6;k++){
		        	String pass="passwrd"+k;
		        	System.out.println(pass+" = "+prop.getProperty(pass));
			        driver.findElement(By.id("Passwd")).sendKeys(prop.getProperty(pass));
		        
		        
		        System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
		        
		        //Click Sign In button
		        driver.findElement(By.id("signIn")).click();
		        
		        System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);
		        
		        //find the element
		        if(SDC.finErrorMsgOnPage()){
		        //check sign in fail
		        String errorMsg=null;
		        errorMsg=driver.findElement(By.id("errormsg_0_Passwd")).getText();
		        
		        if(errorMsg==null){
		        	System.out.println("Log in Pass");
		        }else
		        	System.out.println("Log in Fail : "+errorMsg);
		        }else
		        	System.out.println("Error Message Not Displayed");
		        }
		        
		        String pageTitle=driver.getTitle();
		        if(pageTitle.equals(""))
		        	mStatus=true;
		        else
		        	mStatus=false;
		        mStatus=driver.getPageSource().contains("Name");
				System.out.println(mStatus);
		        return mStatus;	
			}
				
			public Boolean finErrorMsgOnPage(){
				
				System.out.println("finErrorMsgOnPage()");
			
				try {
					driver.findElement(By.id("errormsg_0_Passwd"));
				       return true;
				   } catch (NoSuchElementException e) {
				       return false;
				   }
			
		}
			
			public void signOut() throws Exception{
				
				System.out.println("signOut()");
			
				
			}
			
			public void closeBrowser() throws Exception {
				
				System.out.println("closeBrowser()");
			
				System.out.println("Sleeping (5 sec)");
		        Thread.sleep(5000);
		        
		        //Close the browser
		        System.out.println("Closing the browser");
		        driver.close();
			}
			
			public void composeEmail() throws Exception{
				
				System.out.println("composeEmail()");
			
				//click on Compose Button
				System.out.println("Click on Compose Button");
				driver.findElement(By.xpath(".//*[@id=':4d']/div/div")).click();
				
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
		        
				System.out.println("Enter EmailID");
				WebElement currentElement = driver.switchTo().activeElement();
				currentElement.sendKeys("huzaif.0712@gmail.com");
				
				String elementId=currentElement.getAttribute("id");
				System.out.println(elementId);
				
				driver.findElement(By.xpath(".//*[@id=':9u']")).sendKeys("huzaif.0712@gmail.com");
				
				/*
				
				currentElement.sendKeys(Keys.TAB);
				currentElement.sendKeys(Keys.TAB);
				currentElement.sendKeys("SUB: Selenium Demo");
				currentElement.sendKeys(Keys.TAB);
				currentElement.sendKeys("Hi Hello How are you.");
				currentElement.sendKeys(Keys.TAB);
				
				//driver.findElement(By.xpath(".//*[@id=':9u']")).sendKeys("huzaif.0712@gmail.com");
				*/
				
			}
			
			public void googleDrive() throws Exception {
				
				System.out.println("googleDrive()");
			
				System.out.println("Sleeping (2 sec)");
				Thread.sleep(2000);
				
				driver.findElement(By.xpath(".//*[@id='gbwa']/div[1]/a")).click();
				driver.findElement(By.xpath(".//*[@id='gb25']/span[1]")).click();
				driver.findElement(By.xpath(".//*[@id='navpane']/div[2]/div[1]/div/div/div[1]")).click();
				
			}
			
			public Boolean performClicks() throws Exception{
				Boolean mStatus=false;
				System.out.println("performClicks()");
			
				
				System.out.println("Sleeping (2 sec)");
				Thread.sleep(2000);
				
				driver.findElement(By.xpath(".//*[@id='docs-file-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id=':iu']/div/span[1]")).click();
				driver.findElement(By.xpath(".//*[@id=':iu']/div/span[1]")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-edit-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-edit-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-view-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-view-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-insert-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-insert-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-format-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-format-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='trix-data-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='trix-data-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-tools-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-tools-menu']")).click();
				System.out.println("Sleeping (2 sec)");
		        Thread.sleep(2000);
				driver.findElement(By.xpath(".//*[@id='docs-help-menu']")).click();
				driver.findElement(By.xpath(".//*[@id='docs-help-menu']")).click();
				
				System.out.println("Sleeping (2 sec)");
				Thread.sleep(2000);
				
				mStatus=true;
				return mStatus;
			}
			
			public Boolean compareGDriveSheetWithLocalSheet() throws Exception{
				Boolean mStatus=false;
				System.out.println("compareGDriveSheetWithLocalSheet()");
			
				
				//SeleniumDemoClass test=new SeleniumDemoClass();
				
				int 	rowNumberSharedCopy=1,
						columnNumberSharedCopy=1,
						//rowCountSharedCopy=0,
						//columnCountSharedCopy=0, 
						maxRowCountSharedCopy=sharedCopyRowCount, //row numbers start from 0 to max value
						maxColumnCountSharedCopy=sharedCopyColumnCount, //if column number is 7, actual count is 7-2=5
						i=1,
						j=0;
				
				String  cellDataSharedCopy=null,
						greenColor="rgba(255, 0, 0, 1)",
						//redColor="",
						whiteColor="rgba(0, 255, 0, 1)";
				
				while(i<trueCopyRowCount){
			    	   System.out.println("Outer while loop, i : "+i+", j : "+j+"\n");
			    	   rowNumberSharedCopy=1;
			      
			    	   while(rowNumberSharedCopy<maxRowCountSharedCopy)
			    	   {
			        	System.out.println("inner while loop");
				        columnNumberSharedCopy=1;
				        j=0;
				        
				        driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
				        //cellDataSharedCopy=driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberTrueCopy+"']/td["+columnNumberTrueCopy+"]")).getText();
				        cellDataSharedCopy=driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).getText();
				        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" movieNames : "+SDC.readTrueCopy(i,j));
				        
				        
				        if(cellDataSharedCopy.equals(SDC.readTrueCopy(i,j))){
				        	System.out.println("Inside If Condition");
				        	System.out.println("cellDataTrueCopy : "+SDC.readTrueCopy(i,j)+", cellDataSharedCopy : "+cellDataSharedCopy+"\n");
				        	
				        				        driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
										        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" movieNames : "+SDC.readTrueCopy(i,j));

									        	System.out.println("Data Match ");
									        	
										        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
										        driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
										        
										        //Green Color xpath(".//*[@id='jfk-palette-cell-104']/div")
										        driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-104']/div")).click();
												
										        
										      //start while loop to verify column data
										      while(columnNumberSharedCopy<maxColumnCountSharedCopy){
							        			
							        			driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
							        			cellDataSharedCopy=driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).getText();
							        	        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" movieNames : "+SDC.readTrueCopy(i,j));
							        	        

							        	        if(cellDataSharedCopy.equals(SDC.readTrueCopy(i,j))){
										        	        System.out.println("Inside If Condition");
										        	        System.out.println("movieNames : "+SDC.readTrueCopy(i,j)+", cellDataTrueCopy : "+cellDataSharedCopy+"\n");
										        	        	
												        	System.out.println("Data Match ");
												        	
										        	        driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
										        	        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" movieNames : "+SDC.readTrueCopy(i,j));
							        							        				        
										        	        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
										        	        driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
							        							        
										        	        //Green Color xpath(".//*[@id='jfk-palette-cell-104']/div")
										        	        driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-104']/div")).click();
							        	        }else{
										        	System.out.println("movieNames : "+SDC.readTrueCopy(i,j)+", cellDataTrueCopy : "+cellDataSharedCopy+"\n");

										        	System.out.println("Data Not Match ");
										        	
											        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
											        driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
											        
											        //Red Color xpath(".//*[@id='jfk-palette-cell-101']")
											        driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-101']")).click();
									        		}
							        	        
							        	        System.out.println("Column Number Shared Copy : "+columnNumberSharedCopy);
										        columnNumberSharedCopy++;
										        j++;
										      }
				        		rowNumberSharedCopy=maxRowCountSharedCopy;
					        }
					        else{
						        	//if the background color of the cell is green, skip the coloring and SOP this row has already been checked
					        		String backgroundcolorcode=driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).getCssValue("background-color");
					        		
					        		System.out.println("Inside Else Condition background-color: "+backgroundcolorcode);
					        		if(backgroundcolorcode.equals(greenColor)||backgroundcolorcode.equals(whiteColor)){
					        			System.out.println("This row has already Passed :)\n");
					        		}else{
					        			
							        	System.out.println("Data Not Match ");
							        	
						        	System.out.println("movieNames : "+SDC.readTrueCopy(i,j)+", cellDataTrueCopy : "+cellDataSharedCopy+"\n");
						        	
							        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
							        driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
							        
							        //Red Color xpath(".//*[@id='jfk-palette-cell-101']")
							        driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-101']")).click();
					        		}
						        }
					        rowNumberSharedCopy++;
			    	   		} 
			    	   		i++;
			       		}
				mStatus=true;
				System.out.println(mStatus);
		        return mStatus;	
			}

			public void compareLocalSheets() throws Exception{ //Code not working...!!! :(
				System.out.println("compareLocalSheets()");
						
				int 	rowNumberSharedCopy=0,
						columnNumberSharedCopy=0,
						//rowCountSharedCopy=0,
						//columnCountSharedCopy=0, 
						maxRowCountSharedCopy=15, //row numbers start from 0 to max value
						maxColumnCountSharedCopy=15, //if column number is 7, actual count is 7-2=5
						i=0,
						j=0;
				
				String  cellDataSharedCopy=null;
				
				while(i<20){
			    	   System.out.println("Outer while loop, i : "+i+", j : "+j+"\n");
			    	   rowNumberSharedCopy=0;
			      
			    	   while(rowNumberSharedCopy<maxRowCountSharedCopy)
			    	   {
			        	System.out.println("inner while loop");
				        columnNumberSharedCopy=0;
				        j=0;
				        
				        cellDataSharedCopy=SDC.readSharedCopy(i,j);
				        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" cellDataTrueCopy : "+SDC.readTrueCopy(i,j)+" cellDataSharedCopy : "+cellDataSharedCopy);
				        
				        
				        if(cellDataSharedCopy.equals(SDC.readTrueCopy(i,j))){
				        	cellDataSharedCopy=SDC.readSharedCopy(i,j);
				        	System.out.println("Inside If Condition");
				        	System.out.println("cellDataTrueCopy : "+SDC.readTrueCopy(i,j)+", cellDataSharedCopy : "+cellDataSharedCopy+"\n");
				        	
				        				        //driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
										        //System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" movieNames : "+SDC.readTrueCopy(i,j));
										        
										        System.out.println("Data Match ");
										        cellDataSharedCopy=SDC.readSharedCopy(i,j);
										        
										        editSheet(i,j,cellDataSharedCopy,"GREEN");
										        		
							        			//Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
										        //driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
										        
										        //Green Color xpath(".//*[@id='jfk-palette-cell-104']/div")
										        //driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-104']/div")).click();
												
										        
										      //start while loop to verify column data
										      while(columnNumberSharedCopy<maxColumnCountSharedCopy){
							        			
							        			//driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
							        			cellDataSharedCopy=SDC.readSharedCopy(i,j);
							        	        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" cellDataTrueCopy : "+SDC.readTrueCopy(i,j));
							        	        

							        	        if(cellDataSharedCopy.equals(SDC.readTrueCopy(i,j))){
							        	        			cellDataSharedCopy=SDC.readSharedCopy(i,j);
										        	        System.out.println("Inside If Condition");
										        	        System.out.println("cellDataTrueCopy : "+SDC.readTrueCopy(i,j)+", cellDataSharedCopy : "+cellDataSharedCopy+"\n");
										        	        	
										        	        //driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).click();
										        	        System.out.println("Row number = "+rowNumberSharedCopy+" Column Number = "+columnNumberSharedCopy+" cellDataTrueCopy : "+SDC.readTrueCopy(i,j));
							        							        				        
										        	        System.out.println("Data Match ");
										        	        
										        	        editSheet(i,j,cellDataSharedCopy,"GREEN");
										        	        
										        	        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
										        	        //driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
							        							        
										        	        //Green Color xpath(".//*[@id='jfk-palette-cell-104']/div")
										        	        //driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-104']/div")).click();
							        	        }else{
							        	        	cellDataSharedCopy=SDC.readSharedCopy(i,j);
										        	System.out.println("cellDataTrueCopy : "+SDC.readTrueCopy(i,j)+", cellDataSharedCopy : "+cellDataSharedCopy+"\n");
										        	
										        	System.out.println("Data Not Match ");
										        	
										        	editSheet(i,j,cellDataSharedCopy,"RED");
										        	
											        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
											        //driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
											        
											        //Red Color xpath(".//*[@id='jfk-palette-cell-101']")
											        //driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-101']")).click();
									        		}
							        	        
							        	        System.out.println("Column Number Shared Copy : "+columnNumberSharedCopy);
										        columnNumberSharedCopy++;
										        j++;
										      }
				        		rowNumberSharedCopy=maxRowCountSharedCopy;
					        }
					        else{
					        	
					        	System.out.println("Data Not Match ");
					        	
					        	/*
						        	//if the background color of the cell is green, skip the coloring and SOP this row has already been checked
					        		String backgroundcolorcode=driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+rowNumberSharedCopy+"']/td["+columnNumberSharedCopy+"]")).getCssValue("background-color");
					        		
					        		System.out.println("Inside Else Condition background-color: "+backgroundcolorcode);
					        		if(backgroundcolorcode.equals(greenColor)||backgroundcolorcode.equals(whiteColor)){
					        			System.out.println("This row has already Passed :)\n");
					        		}else{
						        	System.out.println("movieNames : "+SDC.readTrueCopy(i,j)+", cellDataTrueCopy : "+cellDataSharedCopy+"\n");
						        	
							        //Google Cell Background Color Menu xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")
							        driver.findElement(By.xpath(".//*[@id='t-cell-color']/div/div/div[1]/div")).click();
							        
							        //Red Color xpath(".//*[@id='jfk-palette-cell-101']")
							        driver.findElement(By.xpath(".//*[@id='jfk-palette-cell-101']")).click();
					        		}*/
						        } 
						        
					       rowNumberSharedCopy++;
			    	   		} 
			    	   		i++;
			       		}
			}
			
			public void performKeyboardActionsOnGDriveDoc() throws Exception{

				System.out.println("performKeyboardActionsOnGDriveDoc()");
				
				int x=1,y=1;
				for(;x<35;x++){
					//while(y<6)
					driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+x+"']/td["+y+"]")).sendKeys(Keys.DOWN);
					//driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+x+"']/td["+y+"]")).click();
					
					System.out.println("x : "+x+", y : "+y);
				
				}
				//driver.findElement(By.xpath(".//*[@id='0-r-column-head-section$"+x+"']/td["+y+"]")).click();
				
			}

			public void setTrueFilePath(String trueFile) {
				
				System.out.println("setTrueFilePath()");
			
				this.trueFile = trueFile;
			  }

			public String readTrueCopy(int i,int j) throws IOException  {
				
				System.out.println("readTrueCopy()");
				
			    File inputTrueCopy = new File(trueFile);
			    Workbook wt;
			    int rowNumber=i,columnNumber=j;
			    String cellContentsTrue="cellContentsTrue";
			    try {
			    	wt = Workbook.getWorkbook(inputTrueCopy);
			      // Get the first sheet
			      Sheet sheet = wt.getSheet(0);
			      
			      	//Goto the specific cell according to the row and column number combination
			          Cell cellTrue = sheet.getCell(columnNumber, rowNumber);
			          cellContentsTrue=cellTrue.getContents();

			      //System.out.println("true copy-column number : "+columnNumber+" , true copy-row number "+rowNumber+" cellContentsTrue : "+cellContentsTrue);

			      wt.close();
			    } catch (BiffException e) {
			      e.printStackTrace();
			    }
			    return cellContentsTrue; 
			  } 
			
			public void setSharedFilePath(String sharedFile) {
						
						System.out.println("setSharedFilePath()");
					
						this.sharedFile = sharedFile;
					  }

			public String readSharedCopy(int i,int j) throws IOException  {
				
				System.out.println("readSharedCopy()");
				
			    File inputSharedCopy = new File(sharedFile);
			    Workbook ws;
			    int rowNumber=i,columnNumber=j;
			    String cellContentsShared="cellContentsShared";
			    try {
			    	ws = Workbook.getWorkbook(inputSharedCopy);
			      // Get the first sheet
			      Sheet sheet = ws.getSheet(0);
			      
			      	//Goto the specific cell according to the row and column number combination
			          Cell cellShared = sheet.getCell(columnNumber, rowNumber);
			          cellContentsShared=cellShared.getContents();

			          //System.out.println("shared copy-column number : "+columnNumber+" , shared copy-row number "+rowNumber+" cellContentsShared : "+cellContentsShared);
			      ws.close();

			    } catch (BiffException e) {
			      e.printStackTrace();
			    }
			    return cellContentsShared; 
			  } 
			
			public void editSheet(int rn, int cn, String CellData, String colr) throws Exception{

				System.out.println("editSheet()");
				int rowNum=rn;
				int ColNum=cn;			
				
				Workbook workbook1 = Workbook.getWorkbook(new File("D:\\DEV\\Resource_Files\\SharedCopy.xls"));
				WritableWorkbook copy = Workbook.createWorkbook(new File("D:\\DEV\\Resource_Files\\SharedCopy.xls"), workbook1);
				
				WritableSheet sheet0 = copy.getSheet(0); 
				
				if(colr.equals("GREEN")){
					sheet0.addCell(new jxl.write.Label(rowNum, ColNum, CellData, createFormatCellStatus("GREEN")));    //green color
				}else{
					sheet0.addCell(new jxl.write.Label(rowNum, ColNum, CellData, createFormatCellStatus("RED")));    //red color
				}					
				
				copy.write(); 
				copy.close();
				workbook1.close();
			
			}
			
			public WritableCellFormat createFormatCellStatus(String toDo) throws WriteException{
				
				WritableCellFormat result = new WritableCellFormat();
				
				switch (toDo){
					
				case "GREEN":
					result.setBackground(Colour.GREEN);
				    result.setBorder(Border.ALL, BorderLineStyle.THIN);
				    break;
					
				case "RED":
					result.setBackground(Colour.RED);
				    result.setBorder(Border.ALL, BorderLineStyle.THIN);
				    break;
					
				case "BORDER":
					result.setBorder(Border.ALL, BorderLineStyle.THIN);
				    break;
					
				default:
					break;
					
				
				}

				/*
				  if(b == true){
				 
				    result.setBackground(Colour.GREEN);
				    result.setBorder(Border.ALL, BorderLineStyle.THIN);
			    }else{
				    result.setBackground(Colour.RED);
				    result.setBorder(Border.ALL, BorderLineStyle.THIN);
			    }
					*/
			return result;
			}
			
			public void callEditSheet() throws Exception{
				String []a={"A","B","C","D"};

				for(int i=0;i<5;i++){
					for(int j=0;j<5;j++){
					SDC.editSheet(i, j, a[i],"GREEN");
					}
				}
			}

			public void setTestCaseFilePath(String testCaseFile) {
				
				System.out.println("setTestCaseFilePath()");
			
				this.testCaseFile = testCaseFile;
			  }
			
			public String readTestCaseSheet(int i,int j) throws IOException  {
				
				System.out.println("readTestCaseSheet()");
				
			    File inputTC = new File(testCaseFile);
			    Workbook ws;
			    int rowNumber=i,columnNumber=j;
			    String cellContentsTC="cellContentsTC";
			    try {
			    	ws = Workbook.getWorkbook(inputTC);
			      // Get the first sheet
			      Sheet sheet = ws.getSheet(0);
			      
			      	//Goto the specific cell according to the row and column number combination
			          Cell cellShared = sheet.getCell(columnNumber, rowNumber);
			          cellContentsTC=cellShared.getContents();

			          //System.out.println("shared copy-column number : "+columnNumber+" , shared copy-row number "+rowNumber+" cellContentsShared : "+cellContentsShared);
			      ws.close();

			    } catch (BiffException e) {
			      e.printStackTrace();
			    }
			    return cellContentsTC; 
			  } 

			public void setTestCaseFilePathTemp(String testCaseFileTemp) {
				
				System.out.println("setTestCaseFilePath()");
			
				this.testCaseFileTemp = testCaseFileTemp;
			  }
			
			public void setStatusToTestCaseSheet(Boolean setStatus, int rowNum) throws Exception {
				 System.out.println("setStatusToTestCaseSheet()");
				 String Status="Pass";
				 //Label label;
				 //setStatus=true;	          
				 
				 //DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm");
				 Date date = new Date();
				 String currentDate=date.toString();
				 //System.out.println(dateFormat.format(date)); //Thu Sep 04 00:59:10 IST 2014
		        
				Workbook workbook = Workbook.getWorkbook(new File(testCaseFile));
				WritableWorkbook writableWorkbook = Workbook.createWorkbook(new File(testCaseFileTemp), workbook);
				
				
				WritableSheet writableSheet = writableWorkbook.getSheet(0); 
				
				//WritableCell cell = writableSheet.getWritableCell(rowNum, 7); 
				//WritableCell cell2= writableSheet.getWritableCell(rowNum,6);
				System.out.println(rowNum);
				
				//System.out.println(cell.toString());
				if(setStatus){
					Status="PASS";
					
					//Create Cells with contents of different data types.
		            //Also specify the Cell coordinates in the constructor
				//label = new Label(6,rowNum,currentDate);
				writableSheet.addCell(new jxl.write.Label(6, rowNum, currentDate, createFormatCellStatus("BORDER")));
				writableSheet.addCell(new jxl.write.Label(7, rowNum, Status, createFormatCellStatus("GREEN")));
		            
		            //Add the created Cells to the sheet
		          // writableSheet.addCell(label);	            
		            //Write and close the workbook
		            writableWorkbook.write();
				}
					else{	
						Status="FAIL";
						
						//Create Cells with contents of different data types.
			            //Also specify the Cell coordinates in the constructor
					//label = new Label(7,rowNum, "FAIL");
						
						//label = new Label(6,rowNum,currentDate);

						writableSheet.addCell(new jxl.write.Label(6, rowNum, currentDate, createFormatCellStatus("BORDER")));
						writableSheet.addCell(new jxl.write.Label(7, rowNum, Status, createFormatCellStatus("RED")));

						//Add the created Cells to the sheet
			        //writableSheet.addCell(label);	            
			            //Write and close the workbook
			            writableWorkbook.write();
				}
				workbook.close();
				writableWorkbook.close();
				
				
				SDC.delNmov();
				
			}
			
			public void delNmov() throws Exception{
				System.out.println("delNmov()");
				
				//File testCase=new File(testCaseFile);
							
				Workbook workbook1 = Workbook.getWorkbook(new File(testCaseFileTemp));
				WritableWorkbook writableWorkbook1 = Workbook.createWorkbook(new File(testCaseFile), workbook1);
	            writableWorkbook1.write();
	            workbook1.close();
				writableWorkbook1.close();
				
				//File tempTestCase=new File (testCaseFileTemp);
				//tempTestCase.delete();
			}
			
			public void saveScreenshot() throws IOException {
		        File scrFile = ((TakesScreenshot) driver).getScreenshotAs(OutputType.FILE);
		        FileUtils.copyFile(scrFile, new File("./..."));//give file path here.
		    }

	


}
