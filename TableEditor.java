package ru.mai;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import java.util.Scanner;

/**
 * Work with new table
 */
public class TableEditor {

    /**
     * Scanner object
     */
    private final Scanner num = new Scanner(System.in);

    /**
     * Creating a new table
     */
    public void createTable(XWPFDocument document) {
        XWPFParagraph tableParagraph = document.createParagraph();
        tableParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun tableParagraphConfig = tableParagraph.createRun();
        System.out.print("Введите название таблицы: ");
        tableParagraphConfig.setText(num.nextLine());
        System.out.print("Введите число столбоцов у таблицы: ");
        int columnCount = sizeTest();
        System.out.print("Введите число строк у таблицы: ");
        int rowCount = sizeTest();
        num.nextLine();
        XWPFTable table = document.createTable();
        XWPFTableRow tableRow = table.getRow(0);
        System.out.print("Заполнение 1 строки, введите 1 элемент: ");
        tableRow.getCell(0).setText(num.nextLine());
        for (int i = 1; i < columnCount; i++) {
            System.out.print("Заполнение 1 строки, введите " + (i + 1) + " элемент: ");
            tableRow.addNewTableCell().setText(num.nextLine());
        }
        for (int i = 1; i < rowCount; i++) {
            createCell(table, i, columnCount);
        }
    }

    /**
     * Create new cells
     * @param table created table
     * @param i number of row
     * @param columnCount count of column
     */
    private void createCell(XWPFTable table, int i, int columnCount) {
        XWPFTableRow tableRow = table.createRow();
        for (int j = 0; j < columnCount; j++) {
            System.out.print("Заполнение " + (i + 1) + " строки, введите " + (j + 1) + " элемент: ");
            tableRow.getCell(j).setText(num.nextLine());
        }
    }

    /**
     * Checking the correctness of the number of table rows and columns
     */
    private int sizeTest() {
        int size;
        do {
            while (!num.hasNextInt()) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                num.next();
            }
            size = num.nextInt();
            if (size <= 0) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
            }
        } while (size <= 0);
        return size;
    }
}