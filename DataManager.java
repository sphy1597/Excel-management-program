package main;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataManager {

	ArrayList<String> datas = new ArrayList<String>(); // 각 행에대한 내용을 저장

	File file;

	FileInputStream fis;
	XSSFWorkbook workbook;
	XSSFSheet sheet;

	int cells;
	int rows;

	DataManager() { // 엑셀 파일을 객체로 생성

		try {

			file = new File("data.xlsx");
			fis = new FileInputStream(file);
			workbook = new XSSFWorkbook(fis);
			sheet = workbook.getSheetAt(0);

		} catch (Exception e) {
			e.getStackTrace();
		}
	}

	public void ReadData() { // 엑셀에서 값을 읽어옴

		datas.clear();

		int rowindex = 0;
		int colindex = 0;

		// 행수를 가져옴
		rows = sheet.getPhysicalNumberOfRows();

		// 가로로 한줄씩 쭉 가져옴
		for (rowindex = 0; rowindex < rows; rowindex++) {

			// 행읽기
			XSSFRow row = sheet.getRow(rowindex);

			// 셀의 수
			cells = row.getPhysicalNumberOfCells();

			// 한칸씩 가져옴
			for (colindex = 0; colindex < cells; colindex++) {

				XSSFCell cell = row.getCell(colindex);
				String value = "";

				if (cell == null) {
					value = "null";
				} else {

					// 가져온 셀의 종류에 따라 값을 읽어와 저장
					switch (cell.getCellType()) {
					case FORMULA:
						value = cell.getCellFormula();
						break;
					case STRING:
						value = cell.getStringCellValue() + "";
						break;
					case NUMERIC:
						value = cell.getNumericCellValue() + "";
						break;
					case BLANK:
						value = cell.getBooleanCellValue() + "";
						break;
					case ERROR:
						value = cell.getErrorCellValue() + "";
						break;
					}// switch

				} // if-else

				datas.add(value);

			} // for2

		} // for1
	}

	public void PrintData() { // 읽어온 데이터를 프린트
		System.out.println(rows);
		System.out.println(cells);
		System.out.println(datas.get(7));
		for (int i = 0; i < rows * cells; i++) {
			System.out.print(datas.get(i) + "  ");
			if (i % cells == cells - 1) {
				System.out.println("");
			}
		}

	}

	public void WriteData() { // 엑셀 파일에 값을 입력

		XSSFRow row2 = sheet.createRow(0); // 행 >> 몇번째 줄
		XSSFCell cell2 = row2.createCell(0); // 셀 >> 몇번쨰 칸
		int index = 0; // 리스트에서의 인덱스

		for (int i = 0; i < rows; i++) {
			row2 = sheet.createRow(i);
			for (int j = 0; j < cells; j++) {
				cell2 = row2.createCell(j);
				cell2.setCellValue(datas.get(index));
				index++;
			}

		}

		try {

			FileOutputStream fOutStream = new FileOutputStream("data.xlsx");
			workbook.write(fOutStream);
			fOutStream.close();
			System.out.println("엑셀파일생성성공");

		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("엑셀파일생성실패");
			System.out.println("DataManager >> WriteData");
		}

	}

	// x행 줄, y가 셀 칸
	public void changevalue(int _x, int _y, String _value) {

		int index = (_x * cells) - (cells - _y);
		datas.set(index, _value);

	}

}
