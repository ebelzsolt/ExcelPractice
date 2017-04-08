package excelpractice;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;
import java.util.Vector;
import javax.swing.JDialog;
import javax.swing.JFrame;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.ScrollPaneConstants;
import javax.swing.table.DefaultTableModel;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Ébel Zsolt
 */
public class ExcelToJTableGUI extends JFrame {

    private final File file;
    private FileInputStream fileInputStream;    
    private DefaultTableModel tableModel;

    public ExcelToJTableGUI(File file) {
        this.file = file;
        setTitle("Import data from Excel file to JTable");
        setMinimumSize(new Dimension(500, 300));
        setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
        setLocationRelativeTo(null);        
        
        try {
            fillTableWithData();
        } catch (FileNotFoundException ex) {
            System.out.println("File not found");
        } catch (IOException ex) {
            System.out.println("IOException");
        }
    }

    /**
     * Workbook is a high level representation of an Excel workbook.
     *
     * HSSF: denotes the API is for working with Excel 2003 and earlier.
     * XSSF: denotes the API is for working with Excel 2007 and later.
     *
     * @return workbook reference depending on excel file format
     * @throws IOException
     */
    private Workbook getWorkbook() throws IOException {
        Workbook workbook = null;
        if (file.getName().endsWith(".xlsx")) {
            workbook = new XSSFWorkbook(fileInputStream);
        } else if (file.getName().endsWith(".xls")) {
            workbook = new HSSFWorkbook(fileInputStream);
        }
        return workbook;
    }

    /**
     * Read excel file line by line and fill table with data
     *
     * @throws FileNotFoundException
     * @throws IOException
     */
    private void fillTableWithData() throws FileNotFoundException, IOException {
        fileInputStream = new FileInputStream(file);
        Workbook workbook = getWorkbook();
        Sheet firstSheet = workbook.getSheetAt(0);

        int rowCount = 0;

        Iterator<Row> rowIterator = firstSheet.iterator();
        while (rowIterator.hasNext()) {
            Row nextRow = rowIterator.next();
            Vector row = new Vector();

            Iterator<Cell> cellIterator = nextRow.iterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();

                switch (cell.getCellTypeEnum()) {
                    case STRING:
                        row.addElement(cell.getStringCellValue());
                        break;
                    case BOOLEAN:
                        row.addElement(cell.getBooleanCellValue());
                        break;
                    case NUMERIC:
                        row.addElement(cell.getNumericCellValue());
                        break;
                    case BLANK:
                        row.addElement("");
                        break;
                    case ERROR:
                        // Implementation goes here
                        break;
                    case FORMULA:
                        // Implementation goes here
                        break;
                    case _NONE:
                        // Implementation goes here
                        break;
                }
            }
            
            /** JTable implementation */
            if (rowCount == 0) {                
                String[] columnNames = new String[row.size()];                
                tableModel = new DefaultTableModel(columnNames, 0);
                
                JTable table = new JTable(tableModel);
                
                JScrollPane scroll = new JScrollPane(table);
                scroll.setVerticalScrollBarPolicy(
                        ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS);
                scroll.setHorizontalScrollBarPolicy(
                        ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);
                getContentPane().add(BorderLayout.CENTER, scroll);                
            } 
            tableModel.addRow(row);            
            rowCount++;
        }
    }
}
