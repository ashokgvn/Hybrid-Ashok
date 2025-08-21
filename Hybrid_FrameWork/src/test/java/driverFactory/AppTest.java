package driverFactory;

import org.openqa.selenium.WebDriver;
import org.testng.annotations.Test;

import com.relevantcodes.extentreports.ExtentReports;
import com.relevantcodes.extentreports.ExtentTest;
import com.relevantcodes.extentreports.LogStatus;

import CommonFunctions.FunctionLibrary;
import utils.ExcelFileUtil;
public class AppTest {
	String fileinputpath = "./DataTables/DatasEngine.xlsx";
	String fileoutputpath = "./DataTables/HybridResults.xlsx";
	ExtentReports reports;
	ExtentTest logger;
	String TCSheet = "MasterTestCases";
	public static WebDriver driver;
	@Test
	public void startTest()throws Throwable
	{
		String Module_Pass="";
		String Module_Fail="";
		//create object for ExcelfileUtil Class
		ExcelFileUtil xl = new ExcelFileUtil(fileinputpath);
		//count rows in TCSheet
		int rc = xl.getRowCount(TCSheet);
		//iterate all rows in TCSheet
		for(int i=1;i<=rc;i++)
		{
			if(xl.getCellData(TCSheet, i, 2).equalsIgnoreCase("Y"))
			{
				//read testcase from TCSheet
				String TCModule = xl.getCellData(TCSheet, i, 1);
				//define html path
				reports = new ExtentReports("./taget/ExtentReports/"+TCModule+"   "+FunctionLibrary.generateDate()+"  "+".html");
				logger = reports.startTest(TCModule);
				//iterate all rows in TCModule
				for(int j=1;j<=xl.getRowCount(TCModule);j++)
				{
					//call cells from TCModule
					String Description = xl.getCellData(TCModule, j, 0);
					String Object_Type = xl.getCellData(TCModule, j, 1);
					String LType = xl.getCellData(TCModule, j, 2);
					String LValue = xl.getCellData(TCModule, j, 3);
					String TData = xl.getCellData(TCModule, j, 4);
					try {
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver = FunctionLibrary.startBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("openUrl"))
						{
							FunctionLibrary.openUrl();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("waitForElement"))
						{
							FunctionLibrary.waitForElement(LType, LValue, TData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(LType, LValue, TData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(TData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("dropDownAction"))
						{
							FunctionLibrary.dropDownAction(LType, LValue, TData);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("captureStock"))
						{
							FunctionLibrary.captureStock(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("stockTable"))
						{
							FunctionLibrary.stockTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("captureSupplier"))
						{
							FunctionLibrary.captureSupplier(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("supplierTable"))
						{
							FunctionLibrary.supplierTable();
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("captureCustomer"))
						{
							FunctionLibrary.captureCustomer(LType, LValue);
							logger.log(LogStatus.INFO, Description);
						}
						if(Object_Type.equalsIgnoreCase("customerTable"))
						{
							FunctionLibrary.customerTable();
							logger.log(LogStatus.INFO, Description);
						}
						//write as pass into TCModule sheet for each method
						xl.setCelldata(TCModule, j, 5, "Pass", fileoutputpath);
						logger.log(LogStatus.PASS, Description);
						Module_Pass="True";	
					} catch (Exception e) {
						System.out.println(e.getMessage());
						//write as Fail into TCModule sheet for each method
						xl.setCelldata(TCModule, j, 5, "Fail", fileoutputpath);
						logger.log(LogStatus.FAIL, Description);
						Module_Fail="False";
					}
					if(Module_Pass.equalsIgnoreCase("True"))
					{
						//Write as Pass into TCSheet
						xl.setCelldata(TCSheet, i, 3, "Pass", fileoutputpath);
					}
					if(Module_Fail.equalsIgnoreCase("False"))
					{
						xl.setCelldata(TCSheet, i, 3, "Fail", fileoutputpath);
					}
					reports.endTest(logger);
					reports.flush();
				}
			}
			else
			{
				//which testcase flag to N write as Blocked in to TCSheet
				xl.setCelldata(TCSheet, i, 3, "Blocked", fileoutputpath);
			}
		}
	}
}
