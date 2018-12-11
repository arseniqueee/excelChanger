import org.apache.commons.lang3.math.NumberUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
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
import java.text.SimpleDateFormat;
import java.util.Date;

public class ExcelGUI extends JFrame {

    String[] items = {

            "Cian"

    };
    private JComboBox combo = new JComboBox(items);

    private JButton file1 = new JButton("Check");
    private JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
    private JLabel label1 = new JLabel("Исходный файл:");
    private JLabel label2 = new JLabel("Входящий файл:");
    private JLabel label3 = new JLabel();
    private JTextField text1 = new JTextField();
    private JTextField text2 = new JTextField();
    private JButton choose1 = new JButton("Выбрать");
    private JButton choose2 = new JButton("Выбрать");
    private JProgressBar bar = new JProgressBar();


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
        text1.setEditable(false);
        container.add(text2);
        text2.setBounds(115, 40, 300, 20);
        text2.setEditable(false);
        container.add(choose1);
        choose1.setBounds(420, 10, 100, 20);
        container.add(choose2);
        choose2.setBounds(420, 40, 100, 20);
        container.add(label3);
        label3.setBounds(150, 95, 300, 15);
        container.add(bar);
        bar.setBounds(140, 71, 300, 20);
        bar.setStringPainted(true);
        bar.setMinimum(0);
        bar.setMaximum(100);
        container.add(combo);
        combo.setBounds(10, 95, 100, 20);


        file1.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                SwingWorker<String, Void> worker = new SwingWorker<String, Void>() {
                    @Override
                    protected String doInBackground() throws Exception {
                        try {
                            file1.setEnabled(false);
                            if (combo.getSelectedItem() == "Cian") {
                                cian(text1.getText(), text2.getText());
                            }
                            file1.setEnabled(true);
                            return "Done";
                        } catch (Exception ex) {
                            System.out.println(ex.getMessage());
                        }
                        return "Done";
                    }
                };
//                test(text1.getText(), text2.getText());
                worker.execute();
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


    public void cian(String dir1, String dir2) {
        int rowCount = 0;
        String[] textMain = new String[300];
        String[] address = new String[300];
        String[] mr = new String[300];
        String[] np = new String[300];
        String[] street = new String[300];
        String[] non = new String[300];
        String[] vzhp = new String[300];
        String[] snt = new String[300];
        String[] type = new String[300];
        String[] naznach = new String[300];
        double[] square = new double[300];
        double[] squareLife = new double[300];
        Integer[] price = new Integer[300];
        Float[] udel = new Float[300];
        String[] electro = new String[300];
        String[] links = new String[300];
        String[] floor = new String[300];
        String[] allFloor = new String[300];
        String[] material = new String[300];
        String[] houseNum = new String[300];
        bar.setValue(0);
        label3.setText("");
        combo.setEnabled(false);
        choose1.setEnabled(false);
        choose2.setEnabled(false);

        try {
            File myFile = new File(dir1);
            FileInputStream fis = null;
            try {
                fis = new FileInputStream(myFile);
            } catch (FileNotFoundException ex) {
                label3.setText("Первый файл не найден");
            } catch (Exception ex) {
                ex.printStackTrace();
            }
            XSSFWorkbook workbook = new XSSFWorkbook(fis);
            XSSFSheet sheet = workbook.getSheetAt(0);
            XSSFCell cell = null;
            XSSFRow row = null;
            rowCount = sheet.getPhysicalNumberOfRows();
            for (int i = 1; i < rowCount; i++) {
                row = sheet.getRow(i);
                //Описание
                cell = row.getCell(getColumnAdress("Описание", sheet));
                textMain[i] = cell.getStringCellValue();
                //Адрес
                cell = row.getCell(getColumnAdress("Адрес", sheet));
                address[i] = cell.getStringCellValue();
                //Парс МР
                String[] munR = address[i].split(", ");
                if (munR[1].contains("район")) {
                    munR[1] = munR[1].replace("район", "");
                    mr[i] = "р-н " + munR[1];
                } else {
                    mr[i] = "г " + munR[1];
                }
                //Тип
                cell = row.getCell(getColumnAdress("Тип", sheet));
                int dasd = getColumnAdress("Тип", sheet);
                if (cell.getStringCellValue().contains("гараж")) {
                    type[i] = "Гараж";
                } else if (cell.getStringCellValue().contains("квартиры")) {
                    type[i] = "Квартира";
                } else if (cell.getStringCellValue().contains("комнтаы")) {
                    type[i] = "Комната";
                } else if (cell.getStringCellValue().contains("помещения")) {
                    type[i] = "ПСН";
                } else if (cell.getStringCellValue().contains("офиса")) {
                    type[i] = "Офис";
                } else if (cell.getStringCellValue().contains("участка")) {
                    type[i] = "Участок";
                } else if (cell.getStringCellValue().contains("земли")) {
                    type[i] = "Земля";
                }



                //Парс НП
                if (mr[i].contains("р-н")) {
                    if (munR[2].contains("пос.")) {
                        munR[2] = munR[2].replace("пос.", "");
                        np[i] = "п " + munR[2];
                    } else if (munR[2].contains("с.")) {
                        munR[2] = munR[2].replace("с.", "");
                        np[i] = "c " + munR[2];
                    } else if (munR[2].contains("д.")) {
                        munR[2] = munR[2].replace("д.", "д");
                        np[i] = "c " + munR[2];
                    } else if (munR[2].contains("рп")) {
                        munR[2] = munR[2].replace("рп", "");
                        np[i] = "рп " + munR[2];
                    }
                } else {
                    np[i] = mr[i];
                }
                System.out.println(i);
                //Парс улицы
                for (int j = 0; j < munR.length; j++){
                    if (munR[j].contains("улица")){
                        munR[2] = munR[2].replace("улица", "");
                        street[i] = "ул " + munR[2].trim();
                    } else if (munR[j].contains("проезд")) {
                        munR[2] = munR[2].replace("проезд", "");
                        street[i] = "проезд " + munR[2].trim();
                    } else if (munR[j].contains("тупик")) {
                        munR[2] = munR[2].replace("тупик", "");
                        street[i] = "тупик " + munR[2].trim();
                    }
                }
                //Номер дома
                String number = munR[munR.length - 1];
                if (number.matches("[а-яА-я0-9/]+") && number.length() < 8) {
                    houseNum[i] = number;
                }

                //Этажи и материал стен
                if (type[i].equals("Квартира") || type[i].equals("Комната")){
                    cell = row.getCell(getColumnAdress("Дом", sheet));
                    String[] floors = cell.getStringCellValue().split(", ");
                    String[] floorsNew = floors[0].split("/");
                    if (floorsNew[0] != null){
                        floor[i] = floorsNew[0];
                    }
                    if (floorsNew[1] != null){
                        allFloor[i] = floorsNew[1];
                    }
                    if (floors.length > 1){
                        material[i] = floors[1].replace("й","е");
                    }

                } else if (getColumnAdress("Этаж", sheet) != 0){
                    cell = row.getCell(getColumnAdress("Этаж", sheet));
                    String[] floors = cell.getStringCellValue().split("/");
                    floor[i] = floors[0];
                    if (floors.length > 1){
                        allFloor[i] = floors[1];
                    }
                }


//                if (munR.length > 2) {
//                    if (munR[2].contains("улица") || munR[2].contains("проезд") || munR[2].contains("тупик")) {
//                        munR[2] = munR[2].replace("улица", "");
//                        street[i] = "ул " + munR[2].trim();
//                    }
//                } else if (munR.length > 2) {
//                    if (munR[3].contains("улица") || munR[3].contains("проезд") || munR[3].contains("тупик")) {
//                        street[i] = munR[3];
//                    }
//                } else if (munR.length > 3) {
//                    if (munR[4].contains("улица") || munR[4].contains("проезд") || munR[4].contains("тупик")) {
//                        street[i] = munR[4];
//                    }
//                }
                //СНТ
                int columnSNT = 0;
                if ((getColumnAdress("Название коттеджного поселка", sheet) != 0)) {
                    cell = row.getCell(getColumnAdress("Название коттеджного поселка", sheet));
                    snt[i] = cell.getStringCellValue();
                }
                //НОН
                if (type[i].equals("Квартира") || type[i].equals("Комната")) {
                    non[i] = "Жилое помещение";
                } else {
                    non[i] = "Нежилое помещение";
                }
                //ВЖП
                if (type[i] != null) {
                    if (type[i].equals("Квартира") || type[i].equals("Комната"))
                        vzhp[i] = type[i];
                }
                //Назначение здания
                if (type[i].equals("Квартира") || type[i].equals("Комната")){
                    naznach[i] = "Многоквартирный дом";
                } else {
                    naznach[i] = "Нежилое";
                }

                //Площадь
                int columnSquare = 0;
                if (getColumnAdress("Участок", sheet) != 0) {
                    columnSquare = getColumnAdress("Участок", sheet);
                    square[i] = getSquare(columnSquare, row);
                } else if (getColumnAdress("Площадь", sheet) != 0){
                    columnSquare = getColumnAdress("Площадь", sheet);
                    square[i] = getSquare(columnSquare, row);
                } else  {
                    columnSquare = getColumnAdress("Площадь, м2", sheet);
                    cell = row.getCell(columnSquare);
                    String[] parse = cell.getStringCellValue().split("/");
                    if (type[i].equals("Квартира")){
                        if (parse.length > 1){
                            square[i] = Double.parseDouble(parse[0]);
                            squareLife[i] = Double.parseDouble(parse[1]);
                        }else {
                            square[i] = Double.parseDouble(parse[0]);
                            squareLife[i] = Double.parseDouble(parse[0]);
                        }
                    }else if (type[i].equals("Комната")){
                        if (parse.length > 1){
                            square[i] = Double.parseDouble(parse[1]);
                            squareLife[i] = Double.parseDouble(parse[0]);
                        }else {
                            square[i] = Double.parseDouble(parse[0]);
                            squareLife[i] = Double.parseDouble(parse[0]);
                        }
                    }else {
                        square[i] = Double.parseDouble(parse[0]);
                    }


                }


                //Цена
                cell = row.getCell(getColumnAdress("Цена", sheet));
                String[] parse = cell.getStringCellValue().split(" ");
                price[i] = Integer.parseInt(parse[0]);
//                udel[i] = price[i] / square[i];
                //Доп
                cell = row.getCell(getColumnAdress("Дополнительно", sheet));
                electro[i] = cell.getStringCellValue();
                //Ссылки
                cell = row.getCell(getColumnAdress("Ссылка на объявление", sheet));
                links[i] = cell.getStringCellValue();
            }
            bar.setValue(50);
            fis.close();


            File myFileFinal = null;
            try {
                myFileFinal = new File(dir2);
                fis = new FileInputStream(myFileFinal);
            } catch (FileNotFoundException ex) {
                label3.setText("Второй файл не найден");
            } catch (Exception ex) {
                ex.printStackTrace();
                System.out.println(ex.getMessage());
            }
            workbook = new XSSFWorkbook(fis);
            sheet = workbook.getSheetAt(0);
            try {

                for (int i = 1; i < rowCount; i++) {
                    row = sheet.getRow(i);
                    //Описание
                    cell = row.getCell(getColumnAdress("Текст объявления", sheet));
                    cell.setCellValue(textMain[i]);
                    //Назначение
                    if (getColumnAdress("Назначение объекта недвижимости",sheet) != 0){
                        if (non[i] != null){
                            cell = row.getCell(getColumnAdress("Назначение объекта недвижимости", sheet));
                            cell.setCellValue(non[i]);
                        }
                    }
                    //Вид жилого помещ
                    if (getColumnAdress("Вид жилого помещения",sheet) != 0){
                        if (vzhp[i] != null){
                            cell = row.getCell(getColumnAdress("Вид жилого помещения", sheet));
                            cell.setCellValue(vzhp[i]);
                        }
                    }
                    //Назначение здания
                    if (getColumnAdress("Назначение здания, в котором расположено помещение",sheet) != 0){
                        if (naznach[i] != null){
                            cell = row.getCell(getColumnAdress("Назначение здания, в котором расположено помещение", sheet));
                            cell.setCellValue(naznach[i]);
                        }
                    }
                    //Адрес
                    cell = row.getCell(getColumnAdress("Адрес объекта недвижимости (по объявлению)", sheet));
                    cell.setCellValue(address[i]);
                    //Площадь
                    cell = row.getCell(getColumnAdress("Площадь, кв.м. ", sheet));
                    cell.setCellValue(square[i]);
                    if (getColumnAdress("Жилая площадь кв.м.", sheet) != 0){
                        cell = row.getCell(getColumnAdress("Жилая площадь кв.м.", sheet));
                        cell.setCellValue(squareLife[i]);
                    }
                    //Главная ссылка

                    //Материал стен
                    if (getColumnAdress("Материал стен ", sheet) != 0){
                        cell = row.getCell(getColumnAdress("Материал стен ", sheet));
                        cell.setCellValue(material[i]);
                    }
                    //Этажи
                    if (getColumnAdress("Номер этажа", sheet) != 0){
                        cell = row.getCell(getColumnAdress("Номер этажа", sheet));
                        cell.setCellValue(floor[i]);
                    }
                    if (getColumnAdress("Количество этажей здания", sheet) != 0){
                        cell = row.getCell(getColumnAdress("Количество этажей здания", sheet));
                        cell.setCellValue(allFloor[i]);
                    }
                    //Номер дома
                    if (getColumnAdress("Номер дома", sheet) != 0){
                        cell = row.getCell(getColumnAdress("Номер дома", sheet));
                        cell.setCellValue(houseNum[i]);
                    }
                    //Цена
                    cell = row.getCell(getColumnAdress("Полная цена (из источника)", sheet));
                    cell.setCellValue(price[i]);
//                    cell = row.getCell(28);
//                    cell.setCellValue(udel[i]);
                    //МР
                    cell = row.getCell(getColumnAdress("МР или городской округ", sheet));
                    cell.setCellValue(mr[i]);
                    //НП
                    cell = row.getCell(getColumnAdress("НП", sheet));
                    cell.setCellValue(np[i]);
                    //Улица
                    cell = row.getCell(getColumnAdress("Улица", sheet));
                    cell.setCellValue(street[i]);
                    //СНТ
                    cell = row.getCell(getColumnAdress("Наименование СНТ", sheet));
                    cell.setCellValue(snt[i]);
                    //Дата
                    Date date = new Date();
                    SimpleDateFormat formatForDateNow = new SimpleDateFormat("dd.MM.yyyy");
                    cell = row.getCell(getColumnAdress("Дата источника", sheet));
                    cell.setCellValue(date);
                    cell = row.getCell(getColumnAdress("Электричество", sheet));
                    cell.setCellValue("");
                    cell = row.getCell(getColumnAdress("Водоснабжение", sheet));
                    cell.setCellValue("");
                    cell = row.getCell(getColumnAdress("Газоснабжение", sheet));
                    cell.setCellValue("");
                    cell = row.getCell(getColumnAdress("Канализация", sheet));
                    cell.setCellValue("");
                    if (electro[i].contains("Электричество")) {
                        cell = row.getCell(getColumnAdress("Электричество", sheet));
                        cell.setCellValue("да");
                    }
                    if (electro[i].contains("Водоснабжение")) {
                        cell = row.getCell(getColumnAdress("Водоснабжение", sheet));
                        cell.setCellValue("да");
                    }
                    if (electro[i].contains("Газ")) {
                        cell = row.getCell(getColumnAdress("Газоснабжение", sheet));
                        cell.setCellValue("да");
                    }
                    if (electro[i].contains("Канализация")) {
                        cell = row.getCell(getColumnAdress("Канализация", sheet));
                        cell.setCellValue("да");
                    }
                    //Ссылки
                    cell = row.getCell(getColumnAdress("Номер источника (ссылка URL)", sheet));
                    cell.setCellValue(links[i]);
                }

            } catch (NullPointerException ex) {
                System.out.println(ex);
            } catch (Exception ex) {
                ex.printStackTrace();
                System.out.println(ex.getMessage());
            }
            bar.setValue(75);

            try {
                FileOutputStream fileOutputStream = new FileOutputStream(myFileFinal);
                workbook.write(fileOutputStream);
                fis.close();
                fileOutputStream.close();
                label3.setText("Done!");
                bar.setValue(100);
            } catch (FileNotFoundException ex) {
                label3.setText("Закройте файл");
            } catch (Exception ex) {
                System.out.println(ex.getMessage());
            }
            combo.setEnabled(true);
            choose1.setEnabled(true);
            choose2.setEnabled(true);


        } catch (FileNotFoundException e1) {
            System.out.println(e1.getMessage());
        } catch (IOException e1) {
            System.out.println(e1.getMessage());
        } catch (Exception ex) {
            System.out.println(ex.getMessage());
            ex.printStackTrace();
        }
    }


    public int getColumnAdress(String name, Sheet sheet) {
        Row row = sheet.getRow(0);
        int count = 0;
        int columnSize = 0;
        for (Cell c : row) {
            columnSize++;
        }
        for (int i = 0; i < columnSize; i++) {
            Cell c = row.getCell(i);
            if (c.getStringCellValue().equals(name)){
                count = i;
            }
        }
        return count;
    }


    public double getSquare(int columnSquare, Row row) {
        Cell cell = row.getCell(columnSquare);
        String[] parse = cell.getStringCellValue().split(", ");
        if (parse.length == 2) {
            if (NumberUtils.isCreatable(parse[0])) {
                if (parse[1].equals("сот.")) {
                    return Double.parseDouble(parse[0]) * 100;
                } else if (parse[1].equals("га")) {
                    return Double.parseDouble(parse[0]) * 1000;
                } else {
                    return Float.parseFloat(parse[0]);
                }
            } else {
                return Float.valueOf(0);
            }
        } else if (NumberUtils.isCreatable(parse[0])) {
            return Float.parseFloat(parse[0]);
        } else {
            return (float) 0;
        }
    }


    public static void main(String[] args) {
        ExcelGUI gui = new ExcelGUI();
        gui.setVisible(true);
    }
}
