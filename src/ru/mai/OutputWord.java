package ru.mai;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Working with reading a MS Word file
 */
public class OutputWord {

    /**
     * Regular expression waiting for file name input validation
     */
    private final String MASK = "^[A-Za-zА-Яа-яЁё0-9_]+(\\s)?[A-Za-zА-Яа-яЁё0-9_]+$";

    /**
     * Number of menu items
     */
    private final int NUMBER_OF_MENU = 5;

    /**
     * Scanner object, allows you to read data from the console
     */
    private final Scanner num = new Scanner(System.in);

    /**
     * Menu for selecting the ability to work with a MS Word
     */
    public void outputWord() {
        System.out.print("Введите имя файла: ");
        try (FileInputStream fileInputStream = new FileInputStream(nameCheck() + ".docx")) {
            XWPFDocument docxFile = new XWPFDocument(OPCPackage.open(fileInputStream));
            XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(docxFile);
            int number = 0;
            while (number != NUMBER_OF_MENU) {
                displayingStr();
                number = inputCheck();
                switch (number) {
                    case 1:
                        readDocument(docxFile);
                        break;
                    case 2:
                        readAllParagraphs(docxFile);
                        break;
                    case 3:
                        readHeaderModel(headerFooterPolicy);
                        break;
                    case 4:
                        readFooterModel(headerFooterPolicy);
                        break;
                    case 5:
                        break;
                    default:
                        System.out.println("Invalid Data! Вы ввели неверное значение!");
                }
            }
        } catch (FileNotFoundException e) {
            System.out.println("Error! Файл не найден!");
        } catch (Exception e) {
            System.out.println("Error! Ошибка при чтении файла!");
        }
    }

    /**
     * Checking the selection of menu items
     * @return number menu item
     */
    private int inputCheck () {
        int number;
        do {
            while (!num.hasNextInt()) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                num.next();
                displayingStr();
            }
            number = num.nextInt();
            if (number <= 0 || number > NUMBER_OF_MENU) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                displayingStr();
            }
        } while (number <= 0 || number > NUMBER_OF_MENU);
        return number;
    }

    /**
     * Checking the correctness of the entered file name
     */
    private String nameCheck() {
        Pattern pattern;
        Matcher matcher;
        String font;
        do {
            font = num.nextLine();
            pattern = Pattern.compile(MASK);
            matcher = pattern.matcher(font);
            if (!matcher.find()) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                displayingStr();
            }
        } while (matcher.find());
        return font;

    }

    /**
     * List of features for working with the MS Word
     */
    private void displayingStr() {
        System.out.println("Что вы хотите сделать?\n1. Считать весь файл\n2. Считать параграфы\n3. Считать верхний колонтитул");
        System.out.print("4. Считать нижний колонтитул\n5. Назад\nВведите цифру: ");
    }

    /**
     * Reading a header model
     */
    private void readHeaderModel(XWPFHeaderFooterPolicy headerFooterPolicy) {
        XWPFHeader docHeader = headerFooterPolicy.getDefaultHeader();
        System.out.println(docHeader.getText());
    }

    /**
     * Reading a footer model
     */
    private void readFooterModel(XWPFHeaderFooterPolicy headerFooterPolicy) {
        XWPFFooter docFooter = headerFooterPolicy.getDefaultFooter();
        System.out.println(docFooter.getText());
    }

    /**
     * Reading all paragraphs
     */
    private void readAllParagraphs(XWPFDocument docxFile) {
        List<XWPFParagraph> paragraphs = docxFile.getParagraphs();
        int number = 0;
        for (XWPFParagraph p : paragraphs) {
            number++;
            System.out.println("Параграф " + number + ": " + p.getText());
        }
    }

    /**
     * Reading all document
     */
    private void readDocument(XWPFDocument docxFile) {
        XWPFWordExtractor extractor = new XWPFWordExtractor(docxFile);
        System.out.println(extractor.getText());
    }
}