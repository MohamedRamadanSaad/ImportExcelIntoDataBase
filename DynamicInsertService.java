package com.mr.aqaar.services;

import java.io.File;
import java.io.FileInputStream;

import javax.annotation.PostConstruct;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Service;

@Service
public class DynamicInsertService {

	@Autowired
	private JdbcTemplate jdbcTemplate;

	String masterQuery;
	String rowQuery;
	String newST;
	String likeQuery;
	@PostConstruct
	public void init() {
		 masterQuery = "Insert into aqaar.Master_data ( ";
		 likeQuery ="";
		 rowQuery = "";
		 newST="";
	}
	
	
	public void insertFileExcel(String filePath) {
		FileInputStream inputStream = null;
		try {
			inputStream = new FileInputStream(new File(filePath));
			Workbook workbook = new XSSFWorkbook(inputStream);
			Sheet resSheet = workbook.getSheetAt(0);
			XSSFRow rowObj = null;
			
				// if i = 1 then is the header
				rowObj = (XSSFRow) resSheet.getRow(1);
				int numberOfColumnsInRow = rowObj.getPhysicalNumberOfCells();
				for (int column = 1; column < numberOfColumnsInRow; column++) {// from 1 to exclude serial
					Cell cellObj = rowObj.getCell(column);

					// comment row
					
						String comment = cellObj.getCellComment().getString().toString();
						masterQuery += comment + " ";
//						dynamicQuery.append(comment);
						if (column < numberOfColumnsInRow-1 ) {
							masterQuery += " , ";
						
					}
				}

			
			System.out.println(masterQuery);
			
			for (int row = 2; row < resSheet.getPhysicalNumberOfRows(); row++) {
				rowQuery = " ) values ( ";
				// if i = 1 then is the header
				rowObj = (XSSFRow) resSheet.getRow(row);
				for (int column = 1; column < numberOfColumnsInRow; column++) {// from 1 to exclude serial
					DataFormatter formatter = new DataFormatter();
					Cell cellObj = rowObj.getCell(column);

					String cellValue = formatter.formatCellValue(cellObj);
					rowQuery += "'"+cellValue.trim()+"'";
					if (column < numberOfColumnsInRow-1 )
						rowQuery += " , ";
					
					
				}
				
				
//				save hereeeeeeeeeeeee
				rowQuery += " ) ";
				String newST = masterQuery+rowQuery;
				System.out.println("Final Query for row [ "+row+" ] >>> "+newST);
				jdbcTemplate.update(newST);
				newST = "";
			}
			

		} catch (Exception e) {
			e.printStackTrace();
		
			throw new RuntimeException("Error in loading file at server.");

		}finally {
			 masterQuery = "Insert into aqaar.Master_data ( ";
			 rowQuery = "";
			 newST="";
		}

	}

	private void allocateCell(int j, String trim) {
		// TODO Auto-generated method stub

	}
}
