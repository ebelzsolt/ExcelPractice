package excelpractice;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.EventQueue;
import java.awt.Graphics;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.net.URL;
import javax.swing.ImageIcon;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;

/**
 *
 * @author Ébel Zsolt
 */
public class MainGUI extends JFrame {

    JMenuItem open;
    JMenuItem create;
    JMenuItem exit;

    public MainGUI() {
        setTitle("Working with Microsoft Excel files");
        setSize(new Dimension(400, 200));
        setResizable(false);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setJMenuBar(buildMenuBar());
        getContentPane().add(BorderLayout.CENTER, new BackgroundPanel());
    }

    /**
     * Create File menu with three menu items 
     * - New file 
     * - Open file 
     * - Exit
     *
     * @return JMenuBar
     */
    private JMenuBar buildMenuBar() {
        JMenuBar menuBar = new JMenuBar();
        JMenu menu = new JMenu("File");
        create = new JMenuItem("New file");
        open = new JMenuItem("Open file");
        exit = new JMenuItem("Exit");

        create.addActionListener(new MenuListerner());
        open.addActionListener(new MenuListerner());
        exit.addActionListener(new MenuListerner());

        menu.add(create);
        menu.add(open);
        menu.add(exit);
        menuBar.add(menu);
        return menuBar;
    }

    /** Draw background image which is center-aligned */
    private class BackgroundPanel extends JPanel {

        @Override
        protected void paintComponent(Graphics g) {
            try {
                URL url = getClass().getResource("\\img\\java_logo.png");
                Image img = new ImageIcon(url).getImage();
                int x = (getWidth() - img.getWidth(null)) / 2;
                int y = (getHeight() - img.getHeight(null)) / 2;

                g.drawImage(img, x, y,
                        img.getWidth(null),
                        img.getHeight(null),
                        this);
            } catch (NullPointerException ex) {
                System.out.println("Image is missing!");
            }
        }
    }

    /** Custom Excel file filter class for JFileChooser */
    private class ExcelFileFilter extends FileFilter {
        
        @Override
        public boolean accept(File f) {
            if (f.isDirectory()) {
                return true;
            }
            
            String extension = getExtension(f);
            if ((extension != null) &&
                (extension.equals("xls") || extension.equals("xlsx"))) {                
                    return true;            
            }            
            return false;
        }

        @Override
        public String getDescription() {
            return "Excel files (*.xls, *.xlsx)";
        }

        public String getExtension(File f) {
            String ext = null;
            String s = f.getName();
            int i = s.lastIndexOf('.');

            if (i > 0 && i < s.length()- 1) {
                ext = s.substring(i + 1).toLowerCase();
            }
            return ext;
        }
    }

    private class MenuListerner implements ActionListener {

        @Override
        public void actionPerformed(ActionEvent e) {

            if (e.getSource() == open) {
                /** JFileChooser provides a simple mechanism for the user to
                 * choose a file.
                 */
                JFileChooser fileChooser = new JFileChooser();
                
                /** Set up file filters */
                fileChooser.setAcceptAllFileFilterUsed(false);
                fileChooser.setFileFilter(new FileNameExtensionFilter(
                        "Excel 97-2003 Workbook (*.xls)", "xls"));
                fileChooser.setFileFilter(new FileNameExtensionFilter(
                        "Excel Workbook (*.xlsx)", "xlsx"));
                fileChooser.setFileFilter(new ExcelFileFilter());
                fileChooser.showOpenDialog(getContentPane());

                try {
                    File excelFile = fileChooser.getSelectedFile();
                    String fileName = excelFile.getName();

                    //** Validation */
                    if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {
                        EventQueue.invokeLater(new Runnable() {
                            @Override
                            public void run() {
                                new ExcelToJTableGUI(excelFile).setVisible(true);
                            }
                        });
                    } else {
                        JOptionPane.showMessageDialog(
                                null,
                                fileName + " is not an Excel file!\n"
                                + "Please select only Excel file!",
                                "Wrong file",
                                JOptionPane.ERROR_MESSAGE
                        );
                    }
                } catch (NullPointerException ex) {
                    System.out.println("File is not selected");
                }

            } else if (e.getSource() == create) {
                EventQueue.invokeLater(new Runnable() {
                    @Override
                    public void run() {
                        new JTableToExcelGUI().setVisible(true);
                    }
                });

            } else if (e.getSource() == exit) {
                System.exit(0);
            }
        }
    }
}
