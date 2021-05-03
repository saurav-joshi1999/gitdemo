package Api_module.MavenProject;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExemlTest {

	public static void main(String[] args) throws IOException {




	}

	public ArrayList<String> excelMethod(String name) throws FileNotFoundException, IOException {
		ArrayList<String> al = new ArrayList<String>();

		FileInputStream file = new FileInputStream("C:\\Users\\RT_PRIO\\Desktop\\Excel Framework Sele\\Book1.xlsx");
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		int num = workbook.getNumberOfSheets();
		for (int i = 0; i < num; i++) {
			if (workbook.getSheetName(i).equalsIgnoreCase("Sheet1")) {
				XSSFSheet sheet = workbook.getSheetAt(i);

				Iterator<Row> rows = sheet.iterator();
				Row firstrow = rows.next();
				Iterator<Cell> cells = firstrow.cellIterator();
				int k = 0;
				int col = 0;
				while (cells.hasNext()) {
					Cell value = cells.next();
					if (value.getStringCellValue().equalsIgnoreCase("TestCase"))
					{

						col = k;
						// break;
					}
					k++;
				}
				System.out.println(col);
				while (rows.hasNext()) {

					Row rNext = rows.next();

					if (rNext.getCell(col).getStringCellValue().equalsIgnoreCase(name)) {
						Iterator<Cell> cc = rNext.cellIterator();

						while (cc.hasNext()) {
							Cell cv = cc.next();
							if (cv.getCellType() == CellType.STRING) {
							al.add(cv.getStringCellValue());
							} else {
								al.add(NumberToTextConverter.toText(cv.getNumericCellValue()));
							}
						}
					}
				}

		}

	}
	return al;
}


}

