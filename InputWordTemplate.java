package ru.mai;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Working with template MS Word document creation
 */
public class InputWordTemplate {

    /**
     * Create a template document
     */
    public void inputWordTemplate() {
        XWPFDocument document = new XWPFDocument();
        CTSectPr ctSectPr = document.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy headerFooterPolicy = new XWPFHeaderFooterPolicy(document, ctSectPr);

        CTP ctpHeaderModel = createHeaderModel();
        XWPFParagraph headerParagraph = new XWPFParagraph(ctpHeaderModel, document);
        headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT, new XWPFParagraph[]{headerParagraph});

        CTP ctpFooterModel = createFooterModel();
        XWPFParagraph footerParagraph = new XWPFParagraph(ctpFooterModel, document);
        headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT, new XWPFParagraph[]{footerParagraph});

        XWPFParagraph bodyParagraph = document.createParagraph();
        bodyParagraph.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun bodyParagraphConfig = bodyParagraph.createRun();
        bodyParagraphConfig.setItalic(true);
        bodyParagraphConfig.setBold(true);
        bodyParagraphConfig.setFontSize(25);
        bodyParagraphConfig.setColor("1EC738");
        bodyParagraphConfig.setText("Вы создали шаблонный документ, поздравляю!");

        createTemplateTable(document);

        XWPFParagraph infoParagraph = document.createParagraph();
        infoParagraph.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun infoParagraphConfig = infoParagraph.createRun();
        infoParagraphConfig.setFontSize(16);
        infoParagraphConfig.setFontFamily("Times New Roman");
        infoParagraphConfig.setText("Apache POI, a project run by the Apache Software Foundation, and previously a sub-project of the Jakarta Project, ");
        infoParagraphConfig.setText("provides pure Java libraries for reading and writing files in Microsoft Office formats, such as Word, PowerPoint and Excel.");

        XWPFParagraph trInfoP = document.createParagraph();
        trInfoP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun trInfoParConfig= trInfoP.createRun();
        trInfoParConfig.setFontSize(14);
        trInfoParConfig.setFontFamily("Times New Roman");
        trInfoParConfig.setItalic(true);
        trInfoParConfig.setText("Apache POI, проект, запущенный Apache Software Foundation и ранее являвшийся подпроектом Джакартского проекта, предоставляет");
        trInfoParConfig.setText(" чистые библиотеки Java для чтения и записи файлов в форматах Microsoft Office, таких как Word, PowerPoint и Excel.");

        boolean isExceptional = false;
        try {
            FileOutputStream out = new FileOutputStream("Template.docx");
            document.write(out);
            document.close();
        } catch (IOException e) {
            System.out.println("Файл открыт! Пожалуйста закройте программу, которая его использует и повторите попытку!");
            isExceptional = true;
        }
        if (!isExceptional) {
            System.out.println("Файл был создан по шаблону под именем Template.docx!");
        }
    }

    /**
     * Creating template table
     * @param document created document
     */
    private void createTemplateTable(XWPFDocument document) {
        XWPFParagraph tableParagraph = document.createParagraph();
        tableParagraph.setAlignment(ParagraphAlignment.LEFT);
        XWPFRun tableParagraphConfig = tableParagraph.createRun();
        tableParagraphConfig.setText("Таблица 1 - Результаты эксперимента");
        XWPFTable table = document.createTable();
        XWPFTableRow tableRowOne = table.getRow(0);
        tableRowOne.getCell(0).setText("Число строк");
        tableRowOne.addNewTableCell().setText("10 строк");
        tableRowOne.addNewTableCell().setText("100 строк");
        tableRowOne.addNewTableCell().setText("1000 строк");
        XWPFTableRow tableRowTwo = table.createRow();
        tableRowTwo.getCell(0).setText("1");
        tableRowTwo.getCell(1).setText("47,2 мс");
        tableRowTwo.getCell(2).setText("64,7 мс");
        tableRowTwo.getCell(3).setText("136,6 мс");
        XWPFTableRow tableRowThree = table.createRow();
        tableRowThree.getCell(0).setText("2");
        tableRowThree.getCell(1).setText("49,7 мс");
        tableRowThree.getCell(2).setText("77,6 мс");
        tableRowThree.getCell(3).setText("175,9 мс");
        XWPFTableRow tableRowFour = table.createRow();
        tableRowFour.getCell(0).setText("3");
        tableRowFour.getCell(1).setText("43,6 мс");
        tableRowFour.getCell(2).setText("58,1 мс");
        tableRowFour.getCell(3).setText("165,4 мс");
        XWPFTableRow tableRowFive = table.createRow();
        tableRowFive.getCell(0).setText("Сред. знач.");
        tableRowFive.getCell(1).setText("46,8 мс");
        tableRowFive.getCell(2).setText("66,8 мс");
        tableRowFive.getCell(3).setText("159,3 мс");
    }

    /**
     * Creating paragraphs with opportunities library
     * @param document created document
     */
    public void opportunities(XWPFDocument document) {
        XWPFParagraph setCapitalizedP = document.createParagraph();
        setCapitalizedP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setCapitalized = setCapitalizedP.createRun();
        setCapitalized.setCapitalized(true);
        setCapitalized.setText("setCapitalized(boolean value) - Заглавные буквы");

        XWPFParagraph setSmallCapsP = document.createParagraph();
        setSmallCapsP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setSmallCaps = setSmallCapsP.createRun();
        setSmallCaps.setSmallCaps(true);
        setSmallCaps.setText("setSmallCaps(boolean value) - Маленькие заглавные буквы");

        XWPFParagraph setCharacterSpacingP = document.createParagraph();
        setCharacterSpacingP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setCharacterSpacing = setCharacterSpacingP.createRun();
        setCharacterSpacing.setCharacterSpacing(2);
        setCharacterSpacing.setText("setCharacterSpacing(int twips) - Интервал между символами");

        XWPFParagraph setColorP = document.createParagraph();
        setColorP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setColor = setColorP.createRun();
        setColor.setColor("8cfc03");
        setColor.setText("setColor(java.lang.String rgbStr) - Цвет текста");

        XWPFParagraph setDoubleStrikethroughP = document.createParagraph();
        setDoubleStrikethroughP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setDoubleStrikethrough = setDoubleStrikethroughP.createRun();
        setDoubleStrikethrough.setDoubleStrikethrough(true);
        setDoubleStrikethrough.setText("setDoubleStrikethrough(boolean value) - Двойное зачеркивание");

        XWPFParagraph setEmphasisMarkP = document.createParagraph();
        setEmphasisMarkP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setEmphasisMark = setEmphasisMarkP.createRun();
        setEmphasisMark.setText("setEmphasisMark(java.lang.String markType) - Знак подчеркивания");

        XWPFParagraph setShadowP = document.createParagraph();
        setShadowP.setAlignment(ParagraphAlignment.BOTH);
        XWPFRun setShadow = setShadowP.createRun();
        setShadow.setShadow(true);
        setShadow.setText("setShadow(boolean value)  - Тень");
    }

    /**
     * Creating a header model
     * @return header
     */
    private CTP createHeaderModel() {
        CTP ctpHeaderModel = CTP.Factory.newInstance();
        CTR ctrHeaderModel = ctpHeaderModel.addNewR();
        CTText cttHeader = ctrHeaderModel.addNewT();
        cttHeader.setStringValue("Группа: М3О-235Б-19. Студент: Шестихин Н.А.");
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
        cttFooter.setStringValue("Реализация обработки и создания документов MS Word с использованием библиотеки Apache POI");
        return ctpFooterModel;
    }
}