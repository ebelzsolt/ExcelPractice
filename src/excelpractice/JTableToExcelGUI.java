package excelpractice;

import java.awt.BorderLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Vector;
import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;
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
public class JTableToExcelGUI extends JDialog {

    private final String[] columNames = {"First name", "Last name", "Age"};
    private JTextField firstNameTextField;
    private JTextField lastNameTextField;
    private JTextField ageTextField;
    private DefaultTableModel tableModel;
    private JTable table;

    public JTableToExcelGUI() {
        setTitle("JTable to Excel file");
        setResizable(false);
        setModal(true);
        setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);

        getContentPane().add(BorderLayout.NORTH, buildInputPanel());
        getContentPane().add(BorderLayout.CENTER, buildTablePanel());
        getContentPane().add(BorderLayout.SOUTH, createExportButton());

        pack();
        setLocationRelativeTo(null);
    }

    /**
     * Create a JPanel that contains JLabels and JTextfields for user inputs -
     * Three JLabels: First name, Last name and age - Three JTextfields, one for
     * each user input - A JButton for inserting data to JTable
     *
     * @return JPanel
     */
    private JPanel buildInputPanel() {
        /**
         * Grid Layout with 4 rows and 2 columns
         */
        JPanel panel = new JPanel(new GridLayout(4, 2));
        panel.setBorder(BorderFactory.createEmptyBorder(5, 5, 0, 5));
        
        /**
         * JLabel and input field for first name
         */
        panel.add(new JLabel("First name: ", JLabel.RIGHT));
        firstNameTextField = new JTextField(15);
        panel.add(firstNameTextField);

        /**
         * JLabel and input field for last name
         */
        panel.add(new JLabel("Last name: ", JLabel.RIGHT));
        lastNameTextField = new JTextField(15);
        panel.add(lastNameTextField);

        /**
         * JLabel and input field for age
         */
        panel.add(new JLabel("Age: ", JLabel.RIGHT));
        ageTextField = new JTextField(15);
        panel.add(ageTextField);

        /**
         * Empty JLabel to fill the gap so that the "Insert into table" button
         * can be positioned to the right side of the 4th row
         */
        panel.add(new JLabel());

        /**
         * JButton for insert data into the JTable
         */
        JButton insertButton = new JButton("Insert into table");
        insertButton.addActionListener(new InsertButtonListener());
        panel.add(insertButton);
                
        return panel;
    }

    /**
     * Create a JPanel for a JTable Column names are: - First name - Last name -
     * Age
     *
     * @return JPanel containing a JTable
     */
    private JPanel buildTablePanel() {
        JPanel panel = new JPanel();
        tableModel = new DefaultTableModel(columNames, 0);
        table = new JTable(tableModel);
        JScrollPane scrollPane = new JScrollPane(table);

        panel.add(scrollPane);
        return panel;
    }

    /**
     * Create an export button to save content from JTable to an excel file
     *
     * @return exportButton reference
     */
    private JButton createExportButton() {
        JButton exportButton = new JButton("Export");
        exportButton.addActionListener(new ExportButtonListener());
        return exportButton;
    }

    /**
     * Clear all JTextFields
     */
    private void clearInputFields() {
        firstNameTextField.setText("");
        lastNameTextField.setText("");
        ageTextField.setText("");
    }

    private class InsertButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            Vector vector = new Vector();
            String firstName = firstNameTextField.getText();
            String lastName = lastNameTextField.getText();

            vector.add(firstName);
            vector.add(lastName);

            try {
                int age = Integer.valueOf(ageTextField.getText());
                vector.add(age);
            } catch (IllegalArgumentException ex) {
                System.out.println("Hibás életkor!");
            }
            tableModel.addRow(vector);
            clearInputFields();
        }
    }

    /**
     * Returns the selected file from a JFileChooser, including the extension
     * from the file filter.
     */
    public File getSelectedFileWithExtension(JFileChooser c) {
        File file = c.getSelectedFile();
        if (c.getFileFilter() instanceof FileNameExtensionFilter) {
            String[] exts = ((FileNameExtensionFilter) c.getFileFilter()).getExtensions();
            String nameLower = file.getName().toLowerCase();
            for (String ext : exts) { // check if it already has a valid extension
                if (nameLower.endsWith('.' + ext.toLowerCase())) {
                    return file; // if yes, return as-is
                }
            }
            // if not, append the first extension from the selected filter
            file = new File(file.toString() + '.' + exts[0]);
        }
        return file;
    }

    private class ExportButtonListener implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogType(JFileChooser.SAVE_DIALOG);
            fileChooser.setSelectedFile(new File("myExcelFile"));
            fileChooser.setAcceptAllFileFilterUsed(false);
            fileChooser.setFileFilter(new FileNameExtensionFilter(
                    "Excel 97-2003 Workbook (*.xls)", "xls"
            ));
            fileChooser.setFileFilter(new FileNameExtensionFilter(
                    "Excel Workbook (*.xlsx)", "xlsx"
            ));
            fileChooser.showSaveDialog(getContentPane());

            try {
                writeDataToExcelFile(getSelectedFileWithExtension(fileChooser));
            } catch (IOException ex) {
                System.out.println("IOException");
            } catch (NullPointerException ex) {
                System.out.println("Null");
            }

        }
    }

    private void writeDataToExcelFile(File file) throws IOException {
        try (FileOutputStream fout = new FileOutputStream(file)) {

            Workbook wb = null;
            if (file.getName().endsWith("xls")) {
                wb = new HSSFWorkbook();
            } else if (file.getName().endsWith(".xlsx")) {
                wb = new XSSFWorkbook();
            }

            Sheet sheet = wb.createSheet("Sheet 1");

            Row row = sheet.createRow(0);
            Cell cell = null;

            tableModel = (DefaultTableModel) table.getModel();
            for (int i = 0; i < tableModel.getColumnCount(); i++) {
                row.createCell(i).setCellValue(tableModel.getColumnName(i));
            }

            for (int i = 0; i < tableModel.getRowCount(); i++) {
                row = sheet.createRow(i + 1);
                for (int j = 0; j < tableModel.getColumnCount(); j++) {
                    if (tableModel.getValueAt(i, j) != null) {
                        row.createCell(j).setCellValue(
                                tableModel.getValueAt(i, j).toString()
                        );
                    }
                }
            }
            wb.write(fout);

            if (JOptionPane.showConfirmDialog(null, 
                    "Would you like to clear table?", "WARNING",
                    JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION) {
                tableModel.getDataVector().removeAllElements();
                tableModel.fireTableDataChanged();
            }
        }

    }

}
