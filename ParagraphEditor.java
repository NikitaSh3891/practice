package ru.mai;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;

import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Work with new paragraph
 */
public class ParagraphEditor {

    /**
     * Regular expression waiting for font input validation
     */
    private final String MASK = "^[A-Za-z\\s]+$";

    /**
     * Scanner object
     */
    private final Scanner num = new Scanner(System.in);

    /**
     * Creating a new paragraph
     * @param document created document
     */
    public void createParagraph(XWPFDocument document) {
        XWPFParagraph paragraph = document.createParagraph();
        paragraph.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun ParagraphConfig = paragraph.createRun();
        int number = 0;
        while (number != 8) {
            outputStr();
            number = inputCheck(0, 8);
            switch (number) {
                case 1:
                    System.out.print("Введите текст: ");
                    num.nextLine();
                    ParagraphConfig.setText(num.nextLine());
                    break;
                case 2:
                    System.out.print("Введите шкрифт: ");
                    num.nextLine();
                    ParagraphConfig.setFontFamily(fontCheck());
                    break;
                case 3:
                    System.out.print("Введите размер: ");
                    ParagraphConfig.setFontSize(inputCheck(8, 72));
                    break;
                case 4:
                    ParagraphConfig.setBold(true);
                    break;
                case 5:
                    ParagraphConfig.setItalic(true);
                    break;
                case 6:
                    ParagraphConfig.setUnderline(UnderlinePatterns.WORDS);
                    break;
                case 7:
                    ParagraphConfig.setStrikeThrough(true);
                    break;
                case 8:
                    break;
                default:
                    System.out.println("Invalid Data! Вы ввели неверное значение!");
            }
        }
    }

    /**
     * Checking the correctness of the entered font
     */
    private String fontCheck() {
        String font = num.nextLine();
        Pattern pattern = Pattern.compile(MASK);
        Matcher matcher = pattern.matcher(font);
        if (matcher.find()) {
            return font;
        } else {
            System.out.println("Invalid Data! Вы ввели неверное значение!");
            return "Calibri";
        }
    }

    /**
     * Checking the selection of menu items
     * @return number menu item
     */
    private int inputCheck(int lowerBound, int upperBound) {
        int number;
        do {
            while (!num.hasNextInt()) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                num.next();
                outputStr();
            }
            number = num.nextInt();
            if (number <= lowerBound || number > upperBound) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                outputStr();
            }
        } while (number <= lowerBound || number > upperBound);
        return number;
    }

    /**
     * List of features for working with the MS Word
     */
    private void outputStr() {
        System.out.println("Что вы хотите сделать?\n1. Добавить текст\n2. Шрифт\n3. Размер шрифта\n4. Полужирный");
        System.out.print("5. Курсив\n6. Подчеркнутый\n7. Зачеркнутый\n8. Назад\nВведите цифру: ");
    }
}