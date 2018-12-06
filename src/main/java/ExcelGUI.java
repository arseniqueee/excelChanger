import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import javax.swing.filechooser.FileSystemView;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

public class ExcelGUI extends JFrame {

    private JButton file1 = new JButton("Check");
    private JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
    private JLabel label1 = new JLabel("Исходный файл:");
    private JLabel label2 = new JLabel("Входящий файл:");
    private JLabel label3 = new JLabel();
    private JTextField text1 = new JTextField();
    private JTextField text2 = new JTextField();
    private JButton choose1 = new JButton("Выбрать");
    private JButton choose2 = new JButton("Выбрать");


    public ExcelGUI() {
        super("Excel");
        this.setBounds(100, 100, 540, 150);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
        this.setLayout(null);
        this.setResizable(false);


        Container container = this.getContentPane();
        container.add(file1);
        file1.setBounds(10, 70, 100, 20);
        container.add(label1);
        label1.setBounds(10, 10, 100, 15);
        container.add(label2);
        label2.setBounds(10, 40, 100, 15);
        container.add(text1);
        text1.setBounds(115, 10, 300, 20);
//        text1.setEditable(false);
        container.add(text2);
        text2.setBounds(115, 40, 300, 20);
//        text2.setEditable(false);
        container.add(choose1);
        choose1.setBounds(420, 10, 100, 20);
        container.add(choose2);
        choose2.setBounds(420, 40, 100, 20);
        container.add(label3);
        label3.setBounds(150, 75, 300, 15);


        file1.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                test(text1.getText(), text2.getText());
            }
        });

        choose1.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
                int returnValue = jfc.showOpenDialog(null);
                // int returnValue = jfc.showSaveDialog(null);

                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = jfc.getSelectedFile();
                    text1.setText(selectedFile.getAbsolutePath());
                }
            }
        });

        choose2.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
                int returnValue = jfc.showOpenDialog(null);
                // int returnValue = jfc.showSaveDialog(null);

                if (returnValue == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = jfc.getSelectedFile();
                    text2.setText(selectedFile.getAbsolutePath());
                }
            }
        });

    }

    public void test(String dir1, String dir2) {
        try {
            File myFile = new File(dir1);
            FileInputStream fis = null;
            try{
                fis = new FileInputStream(myFile);
            }catch (FileNotFoundException ex){
                label3.setText("Первый файл не найден");
            }
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFCell cell = sheet.getRow(0).getCell(0);
            String str = cell.getStringCellValue();
            fis.close();
            File myFileFinal = null;
            try{
                myFileFinal = new File(dir2);
                fis = new FileInputStream(myFileFinal);
            }catch (FileNotFoundException ex){
                label3.setText("Второй файл не найден");
            }

            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            XSSFRow row = null;
            try {
                row = sheet.getRow(0);
                cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK);
            }catch (NullPointerException ex){
                row = sheet.createRow(0);
                cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK);
            }

            cell.setCellValue(str);
            try{
                FileOutputStream fileOutputStream = new FileOutputStream(myFileFinal);
                workbook.write(fileOutputStream);
                fis.close();
                fileOutputStream.close();
                label3.setText("Done!");
            }catch (FileNotFoundException ex){
                label3.setText("Закройте файл");
            }


        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }


    public static void main(String[] args) {
        ExcelGUI gui = new ExcelGUI();
        gui.setVisible(true);
    }
}
