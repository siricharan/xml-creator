import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

public class ExcelToXMLConverter {

	/**
	 * This method will read given excel <tt>xlsFilePath</tt> file data and
	 * converts into XML format and writes given to <tt>outputFilePath</tt>.
	 * <br>
	 * 
	 * The column names in row '0' in excel shall be treated as tag names and
	 * those corresponding fields in excel will be considered as values in
	 * generated XML file. <br>
	 * 
	 * Each row in excel will be encapsulated within the main tag, which will
	 * have the name as <tt> topTagName </tt>. <br>
	 * 
	 * And all the top tags will be encapsulated within
	 * <tt> topTagName + 's'</tt> string.
	 * 
	 * 
	 * 
	 * @param xlsPath
	 * @param outPath
	 * @param topTagName
	 */
	public void displayFromExcel(String xlsFilePath, String outputFilePath, String topTagName) {

		final File fileSystem = new File(xlsFilePath);

		if (!fileSystem.exists()) {
			System.out.println("File not found in the specified path.");
			return;
		}

		try {
			// Initializing the XML document
			final DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			final DocumentBuilder builder = factory.newDocumentBuilder();
			final Document document = builder.newDocument();
			final Element rootElement = document.createElement(topTagName + "s");
			document.appendChild(rootElement);

			final XSSFWorkbook workBook = new XSSFWorkbook(fileSystem);
			final XSSFSheet sheet = workBook.getSheetAt(0);
			final Iterator<?> rows = sheet.rowIterator();

			final ArrayList<ArrayList<String>> data = new ArrayList<ArrayList<String>>();
			while (rows.hasNext()) {
				final XSSFRow row = (XSSFRow) rows.next();

				final int rowNumber = row.getRowNum();
				if(rowNumber==0 || rowNumber==1)
				{
					continue;
				}
				// display row number
				System.out.println("Row No.: " + rowNumber);

				// get a row, iterate through cells.
				final Iterator<?> cells = row.cellIterator();

				final ArrayList<String> rowData = new ArrayList<String>();
				while (cells.hasNext()) {
					final XSSFCell cell = (XSSFCell) cells.next();
					// System.out.println ("Cell : " + cell.getCellNum ());
					switch (cell.getCellType()) {
					case XSSFCell.CELL_TYPE_NUMERIC: {
						// NUMERIC CELL TYPE
						System.out.println("Numeric: " + cell.getNumericCellValue());
						rowData.add(cell.getNumericCellValue() + "");
						break;
					}
					case HSSFCell.CELL_TYPE_STRING:

					{
						// STRING CELL TYPE
						final XSSFRichTextString richTextString = cell.getRichStringCellValue();

						System.out.println("String: " + richTextString.getString());
						rowData.add(richTextString.getString());
						break;
					}
					default: {
						// types other than String and Numeric.
						System.out.println("Type not supported.");
						break;
					}
					} // end switch

				} // end while
				data.add(rowData);

			} // end while

			workBook.close();

			final int numOfProduct = data.size();

			for (int i = 1; i < numOfProduct; i++) {
				final Element productElement = document.createElement(topTagName);
				rootElement.appendChild(productElement);

				int index = 0;
				for (final String s : data.get(i)) {
					final String headerString = data.get(0).get(index);
					final Element headerElement = document.createElement(headerString);
					productElement.appendChild(headerElement);
					headerElement.appendChild(document.createTextNode(s));
					index++;
				}
			}

			final TransformerFactory tFactory = TransformerFactory.newInstance();

			final Transformer transformer = tFactory.newTransformer();
			// Add indentation to output
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "2");

			final DOMSource source = new DOMSource(document);
			final StreamResult result = new StreamResult(new File(outputFilePath));
			transformer.transform(source, result);

		} catch (final IOException e) {
			System.out.println("IOException " + e.getMessage());
		} catch (final ParserConfigurationException e) {
			System.out.println("ParserConfigurationException " + e.getMessage());
		} catch (final InvalidFormatException e) {
			System.out.println("InvalidFormatException " + e.getMessage());
		} catch (final TransformerConfigurationException e) {
			System.out.println("TransformerConfigurationException " + e.getMessage());
		} catch (final TransformerException e) {
			System.out.println("TransformerException " + e.getMessage());
		}
	}

	/**
	 * @param inputFolder
	 * @return
	 */
	public List<String> getListOfFile(String inputFolder)
	{

		List<String> results = new ArrayList();


		File inputDir = new File(inputFolder);
		if(inputDir.exists() && inputDir.isDirectory())
		{
			File[] files = inputDir.listFiles();
			//If this pathname does not denote a directory, then listFiles() returns null. 

			for (File file : files) {
				if (file.isFile()) {
					results.add(file.getAbsolutePath());
					System.out.println(file.getAbsolutePath());
				}
			}

		}
		
		return results;
	}
	
	
	public static void main(String[] args) {
		final ExcelToXMLConverter poiExample = new ExcelToXMLConverter();
		List<String> listOfExcels = poiExample.getListOfFile(args[0]);
		
		Iterator<String> it = listOfExcels.iterator();
		while(it.hasNext())
		{
			String file = it.next();
			File f = new File(file);
			//String fileNameWithOutExt = FilenameUtils.removeExtension(fileNameWithExt);
			
		     int pos = f.getName().lastIndexOf(".");
			poiExample.displayFromExcel(file, args[1] + "/" + f.getName().substring(0, pos) +".xml", "product");
		}
		
		//
	}
}
