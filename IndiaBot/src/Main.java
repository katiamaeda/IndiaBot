
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;


/*
 * <option value="1">Andaman And Nicobar Islands (UT) </option>
<option value="2">Andhra Pradesh</option>
<option value="3">Arunachal Pradesh</option>
<option value="4">Assam</option>
<option value="5">Bihar</option>
<option value="6">Chandigarh (UT) </option>
<option value="7">Chhattisgarh</option>
<option value="8">Dadra And Nagar Haveli</option>
<option value="9">Daman And Diu (UT) </option>
<option value="10">Delhi (UT) </option>
<option value="11">Goa</option>
<option value="12">Gujarat</option>
<option value="13">Haryana</option>
<option value="14">Himachal Pradesh</option>
<option value="15">Jammu And Kashmir</option>
<option value="16">Jharkhand</option>
<option value="17">Karnataka</option>
<option value="18">Kerala</option>
<option value="19">Lakshadweep (UT) </option>
<option value="20">Madhya Pradesh</option>
<option value="21">Maharashtra</option>
<option value="22">Manipur</option>
<option value="23">Meghalaya</option>
<option value="24">Mizoram</option>
<option value="25">Nagaland</option>
<option value="26">Odisha</option>
<option value="27">Pondicherry (UT) </option>
<option value="28">Punjab</option>
<option value="29">Rajasthan</option>
<option value="30">Sikkim</option>
<option value="31">Tamilnadu</option>
<option value="32">Tripura</option>
<option value="33">Uttar Pradesh</option>
<option value="34">Uttarakhand</option>
<option value="35">West Bengal</option>
<option value="36">Telangana</option>

 * */

public class Main {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		System.setProperty("webdriver.chrome.driver", "/Users/katiamaeda/Downloads/chromedriver");

		int row_num,col_num;
		row_num=1;

		List<Integer> years = Arrays.asList(2015,2014,2013,2012,2011, 2010, 2009, 2008, 2007, 2006, 2005, 2004, 2003, 2002, 2001, 2000);

		List<String> states = Arrays.asList("", "Andaman And Nicobar Islands (UT)", "Andhra Pradesh", "Arunachal Pradesh", "Assam", "Bihar", "Chandigarh (UT)", "Chhattisgarh", "Dadra And Nagar Haveli", 
				"Daman And Diu (UT)", "Delhi (UT)", "Goa", "Gujarat", "Haryana", "Himachal Pradesh", "Jammu And Kashmir", "Jharkhand", "Karnataka", "Kerala", "Lakshadweep (UT)", "Madhya Pradesh",
				"Maharashtra", "Manipur", "Meghalaya", "Mizoram", "Nagaland", "Odisha", "Pondicherry (UT)", "Punjab", "Rajasthan", "Sikkim", "Tamilnadu", "Tripura", "Uttar Pradesh", 
				"Uttarakhand", "West Bengal", "Telangana");

		col_num=1;
		row_num=0;
		
		Row row = null;

		for (int state = 1; state <= 36; state++) {
			System.out.println("starting state = "+state);

			try {
				String fileFolder = "/Users/katiamaeda/Downloads/";
				String inputFile = fileFolder+"Status of Connectivity Total " + (state-1) + ".xlsx";

				SXSSFWorkbook workbook = null;
				Sheet excelSheet = null;
				try {
					FileInputStream inputStream = new FileInputStream(inputFile);
					XSSFWorkbook wb_template = new XSSFWorkbook(inputStream);
					inputStream.close();
					
					workbook = new SXSSFWorkbook(wb_template); 
					excelSheet = workbook.getSheetAt(0);
				} catch (FileNotFoundException e1) {
					XSSFWorkbook wb_template = new XSSFWorkbook();
					workbook = new SXSSFWorkbook(wb_template); 
					excelSheet = workbook.createSheet();
				}

				
				try {
					row = excelSheet.createRow(row_num);
				} catch (Exception e) {
					System.out.println("Exception"+row_num);
				}

				WebDriver driver = new ChromeDriver();
				
//				for (int year = 0; year < years.size(); year++) {
//					System.out.println("starting year = "+years.get(year));

					// And now use this to visit Google
//					driver.get("http://omms.nic.in/MvcReportViewer.aspx?_r=%2fPMGSYCitizen%2fStateWiseListOfWorksSubReport&Level=4&State=" + state + "&District=0&Block=0&Year=" + years.get(year) + "&Batch=0&Collaboration=0&PMGSY=1&Status=%25&LocationName=Andaman+And+Nicobar+Islands+(UT)+&DistrictName=All+Districts&BlockName=All+Blocks&LocalizationValue=en&BatchName=All+Batches&CollaborationName=All+Collaborations&StatusName=All+Status");
					//driver.get("http://omms.nic.in/MvcReportViewer.aspx?_r=%2fPMGSYCitizen%2fUspPropPhysicalProgressofWorksSubreport&Level=4&State=" + state + "&District=0&Block=0&Year=" + years.get(year) + "&Batch=0&Collaboration=0&PMGSY=1&LocationName=Assam&DistrictName=All+Districts&BlockName=All+Blocks&LocalizationValue=en&BatchName=All+Batches&CollaborationName=All+Collaborations");
					//driver.get("http://omms.nic.in/MvcReportViewer.aspx?_r=%2fPMGSYCitizen%2fHabitationCoverage&Level=1&State=" + state + "&District=0&Block=0&PMGSY=1&StateName=Andaman+And+Nicobar+Islands+(UT)+&DistName=All+Districts&BlockName=All+Blocks&LocalizationValue=en");
					driver.get("http://omms.nic.in/StateProfile/StateProfile/SPDetails/" + state + "$0$0$0?_=1438111384177");
					
					
//					try {
//						WebDriverWait wait = new WebDriverWait(driver, 999999999);
//						wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("id('ReportViewer_fixedTable')/tbody/tr[5]/td[3]/div[@id='ReportViewer_ctl09']/div[@id='VisibleReportContentReportViewer_ctl09']/div")));
//					} catch (Exception e) {
//						e.printStackTrace();
//
//						col_num+=5;
//					}

//					WebElement table_element = driver.findElement(By.xpath("id('ReportViewer_fixedTable')/tbody/tr[5]/td[3]/div[@id='ReportViewer_ctl09']/div[@id='VisibleReportContentReportViewer_ctl09']/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[5]/td[3]/table"));

//					WebElement table_element = driver.findElement(By.xpath("id('ReportViewer_fixedTable')/tbody/tr[5]/td[3]/div[@id='ReportViewer_ctl09']/div[@id='VisibleReportContentReportViewer_ctl09']/div/table/tbody/tr/td/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[6]/td[3]/table"));

//					WebElement table_element = driver.findElement(By.xpath("id('divLoadSPDetails')/div[@id='divContentSPConnStat']/div[@id='divContentSPHabCoverage']/table"));
					WebElement table_element;
					try {
						table_element = driver.findElement(By.xpath("id('divLoadSPDetails')/div[@id='divContentSPConnStat']/table"));
					} catch (Exception e) {
						driver.quit();
						System.out.println("finished state = "+state);

						excelSheet.removeRow(row);
						
						String inputFile2 = fileFolder+"/Status of Connectivity Total " + state + ".xlsx";
						FileOutputStream out = new FileOutputStream(inputFile2);
						workbook.write(out);
						out.flush();
					    out.close();
					    
						continue;

					}
					
					List<WebElement> tr_collection=table_element.findElements(By.tagName("tr"));

					int maxColumn = 0;
					for(WebElement trElement : tr_collection) {
						String rowString = trElement.getText();
//						System.out.println("rowString = "+rowString);
						
						rowString = rowString.replaceAll("\n  ", "\n\n");
						String[] separated = rowString.split("\n");
						
						List<WebElement> cols = trElement.findElements(By.tagName("td"));
					    int numberOfColumn = cols.size();
					    if (maxColumn < numberOfColumn) {
					    	maxColumn = numberOfColumn;
					    }

						boolean changed = true;
						if (separated[0].startsWith("Total Habitations Covered")) {
							break;
						}
						if (separated[0].startsWith("New Connectivity") || numberOfColumn <= 1) {
							changed = false;
						}
						
						if (changed) {
							addLabel(col_num-1, row, states.get(state));
							col_num++;
							if (numberOfColumn < maxColumn) {
								col_num = maxColumn-numberOfColumn+2;
						    }
							for (WebElement tdElement : cols) {
								addLabel(col_num-1, row, tdElement.getText());
								col_num++;
							}
						}
						
						
//						for(int i = 0; i < separated.length; i++) {
//							if ((i == 0) && (separated[i].equals("") || separated[i].equals(" ") || separated[i].equals("Sr.No.") || separated[i].equals("Sr. No.") || separated[i].equals("    Total")|| separated[i].equals("                Length")|| separated[i].equals("Total") || separated[i].equals("  Total") || separated[i].equals("    Total No. of Habitations") || separated[i].equals("    1000+"))) {
//								changed = false;
//								break;
//							}
//
//							if (i == 0) {
//								addLabel(col_num-1, row, states.get(state));
//								col_num++;
//							}
//							
//							if (numberOfColumn < maxColumn) {
//								col_num = maxColumn-numberOfColumn+2+i;
//						    }
//
//							addLabel(col_num-1, row, separated[i]);
//							col_num++;
//						}

						if (changed) {
							row_num++;
							row = excelSheet.createRow(row_num);
						}

						col_num=1;
					}

//					System.out.println("finished year = "+years.get(year));
//				}

				driver.quit();
				System.out.println("finished state = "+state);

				excelSheet.removeRow(row);
				
				String inputFile2 = fileFolder+"/Status of Connectivity Total " + state + ".xlsx";
				FileOutputStream out = new FileOutputStream(inputFile2);
				workbook.write(out);
				out.flush();
			    out.close();

			} catch (FileNotFoundException e) {
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		// Check the title of the page
		System.out.println("finished all");

	}

	private static void addLabel(int column, Row row, String s) {

		Cell cell = row.createCell(column);
		cell.setCellValue(s);
	}
}
