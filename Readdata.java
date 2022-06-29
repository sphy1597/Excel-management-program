import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Readdata {
	
	ArrayList<String> values = new ArrayList<String>();
	ArrayList<String> colname = new ArrayList<String>(); // 칼럼이름 저장 
	ArrayList<Dataform> datas = new ArrayList<Dataform>(); // 데이터폼으로 데이터 객체 저장
	
	File file;
	FileInputStream fis;
	XSSFWorkbook workbook;
	
	Readdata(File _file, FileInputStream _fis, XSSFWorkbook _workbook){
		this.file = _file;
		this.fis = _fis;
		this.workbook = _workbook;
	
	
		
	}
	
	
	
	
	public void read() {
		try{
			
			File file = new File("data.xlsx");
			
			FileInputStream fis = new FileInputStream(file);
			XSSFWorkbook workbook = new XSSFWorkbook(fis);
			
			// 세로 1 2 3 행수
			int rowindex = 0;
			// 가로 A B C 셀수
			int colindex = 0;
			
			//시트수
			XSSFSheet sheet = workbook.getSheetAt(0);
			//행 수
			int rows = sheet.getPhysicalNumberOfRows();
			
			//제일 위 항목( 컬럼 ) 가져와서 자장 
			XSSFRow namerow = sheet.getRow(1);
			int namecells = namerow.getPhysicalNumberOfCells();
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
				
			}
			
			//rowindex >> 세로로 몇번째 부터 시작할지 1 2 3 4 
			for(rowindex=1;rowindex<rows;rowindex++) {
				
				//행읽기 
				XSSFRow row = sheet.getRow(rowindex);

				//셀의 수
				int cells = row.getPhysicalNumberOfCells();
				
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
			
			
		}catch(Exception e) {
			e.printStackTrace();
		}
	}//read
	
	
	public void printdata() {
		
		for(int i = 0 ; i < datas.size() ; i ++) {
			System.out.println(i+1+"번");
			System.out.println(datas.get(i).id);
			System.out.println(datas.get(i).name);
			System.out.println(datas.get(i).num);
			System.out.println("==============");			
		}
		
		
	}
}

