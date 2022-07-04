import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataManager {
	
	ArrayList<String> colname = new ArrayList<String>();
	ArrayList<Dataform> datas = new ArrayList<Dataform>();
	ArrayList<String> values = new ArrayList<String>();
	
	
	File file;
	
	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet sheet;
	
	DataManager(){
		
		try {
			
    		file = new File("data.xlsx");
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);		
			sheet = workbook.getSheetAt(0);
			
    	}catch(Exception e) {
    		e.getStackTrace();
    	}
	}
	
	
	public void ReadData() {
		
		datas.clear();
		
		int rowindex = 0;
		int colindex = 0;

		//행수를 가져옴
		int rows = sheet.getPhysicalNumberOfRows();
		
		//칼럼들을 가져옴
		XSSFRow namerow = sheet.getRow(1);		
		int namecells = namerow.getPhysicalNumberOfCells();
		
		//반복문으로 리스트에 저장 
		for(int i=0;i<namecells;i++) {

			XSSFCell namecell = namerow.getCell(colindex);
			String namevalue = "";
			
			if(namecell==null) {
				continue;
			}else {
				switch(namecell.getCellType()) {
    			case FORMULA:
    				namevalue = namecell.getCellFormula();
    				break;
    			case STRING:
    				namevalue = namecell.getStringCellValue()+"";
    				break;
    			case NUMERIC:
    				namevalue = namecell.getNumericCellValue()+"";
    				break;
    			case BLANK:
    				namevalue = namecell.getBooleanCellValue()+"";
    				break;
    			case ERROR:
    				namevalue = namecell.getErrorCellValue()+"";
    				break;
    			}
				
				colname.add(namevalue); // 컬럼이름들 리스트에 추가
			}
			
		}//칼럼 가져오는 반복문
		
		//가로로 한줄씩 쭉 가져옴 
		
		
		for(rowindex=1;rowindex<rows;rowindex++) {
			
			//행읽기 
			XSSFRow row = sheet.getRow(rowindex);

			//셀의 수
			int cells = row.getPhysicalNumberOfCells();
			
			//한칸씩 가져옴 
			for(colindex = 0 ; colindex < cells; colindex++) {
				
				XSSFCell cell = row.getCell(colindex);
				String value = "";
				   				
				
				if(cell==null) {
					continue;
				}else {
					
					//가져온 셀의 종류에 따라 값을 읽어와 저장
	    			switch(cell.getCellType()) {
	    			case FORMULA:
	    				value = cell.getCellFormula();
	    				break;
	    			case STRING:
	    				value = cell.getStringCellValue()+"";
	    				break;
	    			case NUMERIC:
	    				value = cell.getNumericCellValue()+"";
	    				break;
	    			case BLANK:
	    				value = cell.getBooleanCellValue()+"";
	    				break;
	    			case ERROR:
	    				value = cell.getErrorCellValue()+"";
	    				break;
	    			}
	    			
				}
				
				values.add(value);				
			}
			
				Dataform data = new Dataform(values.get(0), values.get(1), values.get(2));
				values.clear();
				datas.add(data);
				
		}
	}
	
	public void PrintData() {
		for(int i = 0 ; i < datas.size() ; i ++) {
			System.out.println(i+1+"번");
			System.out.println(datas.get(i).id);
			System.out.println(datas.get(i).name);
			System.out.println(datas.get(i).num);
			System.out.println("==============");			
		}	
	}
	
	public void WriteData() {
						
		XSSFRow row = sheet.getRow(5);		
		XSSFCell cell = row.getCell(0);
		cell.setCellValue("aaa");
		
	}
	
	
}
