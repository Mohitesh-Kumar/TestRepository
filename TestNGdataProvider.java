package com.jio.testing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.testng.annotations.*;

public class TestNGdataProvider {

	@Test(dataProvider = "getData")
	public void dataTest(String FirstName, String LastName,String Department) {
	System.out.println("FirstName :: "+FirstName);
	System.out.println("LastName :: "+LastName);
	System.out.println("Department :: "+Department);


	}

	@DataProvider(name="getData")
	public Object[][] getData() {
		String[][] obj=new String[3][3];
		String filePath=System.getProperty("user.dir")+"\\src\\main\\resources\\testData.xlsx";
		String sheetName="Sheet1";
		try {
			obj = (String[][]) readDatafromExcel(filePath, sheetName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
return obj;
	
	}

	@SuppressWarnings({ "null", "null" })
	public Object[][] readDatafromExcel(String Filepath, String sheetName) throws IOException {
		

			
			File file = new File(Filepath);
			FileInputStream fis = new FileInputStream(file);
			@SuppressWarnings("resource")
			Workbook workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheet(sheetName);
			int rowCount = sheet.getLastRowNum() - sheet.getFirstRowNum();
			int colCount = sheet.getRow(0).getLastCellNum() - sheet.getRow(0).getFirstCellNum();

String[][] cellValue=new String[rowCount][colCount];
			for (int i = 0; i < rowCount; i++) {

				for (int j = 0;j < colCount; j++) {
					cellValue[i][j] = sheet.getRow(i).getCell(j).toString();
				}

			}
			return cellValue;	

			

	}
}
