import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

public class ExcelGUI extends JFrame {

    private JButton file1 = new JButton("Hi Baby");

    public ExcelGUI(){
        super("Excel");
        this.setBounds(100, 100, 200, 100);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.add(file1);
        file1.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {


                    zhepa();



            }
        });
    }

    public static void zhepa(){
        try {
            File myFile = new File("E://zhepa.xlsx");
            FileInputStream fis = new FileInputStream(myFile);
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFCell cell = sheet.getRow(1).getCell(3);
            cell.setCellValue("Zhepa");
            FileOutputStream stream = new FileOutputStream(myFile.toString());
            workbook.write(stream);
            fis.close();
            stream.close();
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        } catch (IOException e1) {
            e1.printStackTrace();
        }
    }


    public static void main(String[] args){
        ExcelGUI gui = new ExcelGUI();
        gui.setVisible(true);
    }
}
