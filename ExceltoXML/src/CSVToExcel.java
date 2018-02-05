import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.FileReader;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVToExcel {

	public static void main(String[] args) {
		createExcelFromCSV(args[0], args[1]);
	}

	/**
	 * This will convert CSV file to Excel file.
	 * 
	 * @param csvFileAddress
	 * @param xlsxFileAddress
	 */
	private static void createExcelFromCSV(String csvFileAddress, String xlsxFileAddress) {

		try {
			final XSSFWorkbook workBook = new XSSFWorkbook();
			final XSSFSheet sheet = workBook.createSheet("sheet1");
			String currentLine = null;
			int RowNum = 0;
			final BufferedReader br = new BufferedReader(new FileReader(csvFileAddress));
			while ((currentLine = br.readLine()) != null) {
				final String str[] = currentLine.split(",");
				RowNum++;
				final XSSFRow currentRow = sheet.createRow(RowNum);
				for (int i = 0; i < str.length; i++) {
					currentRow.createCell(i).setCellValue(str[i]);
				}
			}

			final FileOutputStream fileOutputStream = new FileOutputStream(xlsxFileAddress);
			workBook.write(fileOutputStream);
			fileOutputStream.close();
			System.out.println("Done");
		} catch (final Exception ex) {
			System.out.println(ex.getMessage() + "Exception in try");
		}

	}
}
