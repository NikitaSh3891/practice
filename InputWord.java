package ru.mai;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTText;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Working with manual file filling (file name, header model, footer model)
 */
public class InputWord {

    /**
     * File name
     */
    private String fileName = "Test";

    /**
     * Regular expression waiting for file name input validation
     */
    private final String MASK = "^[A-Za-zА-Яа-яЁё0-9_]+(\\s)?[A-Za-zА-Яа-яЁё0-9_]+$";

    /**
     * Scanner object
     */
    private final Scanner num = new Scanner(System.in);

    /**
     * Menu for selecting the ability to work with a MS Word
     */
    public void menu() {
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);

        int number = 0;
        while (number != 7 && number != 8) {
            outputStr();
            number = inputCheck();
            switch (number) {
                case 1:
                    nameCheck();
                    break;
                case 2:
                    ParagraphEditor paragraphEditor = new ParagraphEditor();
                    paragraphEditor.createParagraph(document);
                    break;
                case 3:
                    CTP ctpHeaderModel = createHeaderModel();
                    XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
                    headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, new XWPFParagraph[]{headerParagraph});
                    break;
                case 4:
                    CTP ctpFooterModel = createFooterModel();
                    XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooterModel, document);
                    headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, new XWPFParagraph[]{footerParagraph});
                    break;
                case 5:
                    InputWordTemplate inputWordTemplate = new InputWordTemplate();
                    inputWordTemplate.opportunities(document);
                    break;
                case 6:
                    TableEditor tableEditor = new TableEditor();
                    tableEditor.createTable(document);
                    break;
                case 7:
                    mainEditor(document);
                    break;
                case 8:
                    break;
                default:
                    System.out.println("Invalid Data! Вы ввели неверное значение!");
            }
        }
    }

    /**
     * List of features for working with the MS Word
     */
    private void outputStr() {
        System.out.println("Что вы хотите сделать?\n1. Изменить имя файла\n2. Добавить параграф\n3. Редактировать верхний колонтитул\n4. Редактировать нижний колонтитул");
        System.out.print("5. Добавить в файл список некоторых возможностей библиотеки\n6. Добавить таблицу\n7. Сохранить файл.\n8. Назад\nВведите цифру: ");
    }

    /**
     * Saving all the changes you made and writing them to a file
     * @param document created document
     */
    private void mainEditor(XWPFDocument document) {
        fileName += ".docx";
        boolean isExceptional = false;
        try {
            FileOutputStream out = new FileOutputStream(fileName);
            document.write(out);
            document.close();
        } catch (IOException e) {
            System.out.println("Файл открыт! Пожалуйста закройте программу, которая его использует и повторите попытку!");
            isExceptional = true;
        }
        if (!isExceptional) {
            System.out.println("Файл был создан по шаблону под именем " + fileName);
        }
    }

    /**
     * Checking the correctness of the entered file name
     */
    private void nameCheck() {
        System.out.print("Введите название файла: ");
        num.nextLine();
        String name = num.nextLine();

        Pattern pattern = Pattern.compile(MASK);
        Matcher matcher = pattern.matcher(name);
        if (matcher.find()) {
            fileName = name;
            System.out.println("Имя успешно изменено!");
        } else {
            System.out.println("Invalid Data! Вы ввели неверное значение!");
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
                outputStr();
            }
            number = num.nextInt();
            if (number <= 0 || number > 8) {
                System.out.println("Invalid Data! Вы ввели неверное значение!");
                outputStr();
            }
        } while (number <= 0 || number > 8);
        return number;
    }

    /**
     * Creating a header model
     * @return header
     */
    private CTP createHeaderModel() {
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();
        System.out.print("Введите верхний колонтитл: ");
        num.nextLine();
        cttHeader.setStringValue(num.nextLine());
        return ctpHeaderModel;
    }

    /**
     * Creating a footer model
     * @return footer
     */
    private CTP createFooterModel() {
        CTP ctpFooterModel = CTP.Factory.newInstance();
        CTR ctrFooterModel = ctpFooterModel.addNewR();
        CTText cttFooter = ctrFooterModel.addNewT();
        System.out.print("Введите нижний колонтитл: ");
        num.nextLine();
        cttFooter.setStringValue(num.nextLine());
        return ctpFooterModel;
    }
}