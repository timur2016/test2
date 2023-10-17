package tst;

import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main2 {

	public static void main(String[] args) throws Exception {

		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = workbook.createSheet("Avengers");

		Row row1 = sheet.createRow(0);
		row1.createCell(0).setCellValue("IRON-MAN");
		Row row2 = sheet.createRow(1);
		row2.createCell(0).setCellValue("SPIDER-MAN");
		
		row1.setHeightInPoints(75.0f);
		row2.setHeightInPoints(75.0f);

		byte[] spidermanBytes = Main2.class.getClassLoader().getResourceAsStream("spiderman.jpg").readAllBytes();
		byte[] ironmanBytes = Main2.class.getClassLoader().getResourceAsStream("ironman.jpg").readAllBytes();

		int spidermanId = workbook.addPicture(spidermanBytes, Workbook.PICTURE_TYPE_JPEG);
		int ironmanId = workbook.addPicture(ironmanBytes, Workbook.PICTURE_TYPE_JPEG);

		XSSFDrawing drawing = (XSSFDrawing) sheet.createDrawingPatriarch();

		XSSFClientAnchor ironmanAnchor = new XSSFClientAnchor();
		XSSFClientAnchor spidermanAnchor = new XSSFClientAnchor();

		ironmanAnchor.setCol1(1); // Sets the column (0 based) of the first cell.
		ironmanAnchor.setCol2(2); // Sets the column (0 based) of the Second cell.
		ironmanAnchor.setRow1(0); // Sets the row (0 based) of the first cell.
		ironmanAnchor.setRow2(1); // Sets the row (0 based) of the Second cell.

		spidermanAnchor.setCol1(1);
		spidermanAnchor.setCol2(2);
		spidermanAnchor.setRow1(1);
		spidermanAnchor.setRow2(2);

		XSSFPicture pic1 = drawing.createPicture(ironmanAnchor, ironmanId);
		XSSFPicture pic2 = drawing.createPicture(spidermanAnchor, spidermanId);

		pic1.resize();
		pic2.resize();
		for (int i = 0; i < 4; i++) {
			sheet.autoSizeColumn(i);
		}

		try (FileOutputStream saveExcel = new FileOutputStream("test.xlsx")) {
			workbook.write(saveExcel);
		}
		workbook.close();

	}
}
