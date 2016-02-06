package jp.designrule;

import org.junit.Test;
import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.FileOutputStream;
import java.io.FileInputStream;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import jp.designrule.ExcelFormatter;

public class ExcelFormatterTest {
      @Test
      public void testReplaceFormula() {
        File file = new File("sample/sample0.xlsx");
        Workbook workbook = null;
        try(InputStream is = new FileInputStream(file)) {
            workbook = WorkbookFactory.create(is);
          }catch(Exception e) {
          e.printStackTrace();
        }
        Sheet sheet = workbook.getSheet("Sheet1");
        
        ExcelFormatter target = new ExcelFormatter();
        target.replaceFormula(sheet, "B31", "'01_0表紙'!t28");
        
        CellReference reference = new CellReference("B31");
        Row row = sheet.getRow(reference.getRow());
        Cell cell = row.getCell(reference.getCol());        
        assertEquals("'01_0表紙'!t28", cell.getCellFormula());
      }
}
