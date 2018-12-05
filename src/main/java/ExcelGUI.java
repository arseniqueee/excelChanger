import javax.swing.*;
import java.awt.*;

public class ExcelGUI extends JFrame {

    private JButton file1 = new JButton("Hi Baby");

    public ExcelGUI(){
        super("Excel");
        this.setBounds(100, 100, 200, 100);
        this.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);

        Container container = this.getContentPane();
        container.add(file1);
    }

    public static void main(String[] args){
        ExcelGUI gui = new ExcelGUI();
        gui.setVisible(true);
    }
}
