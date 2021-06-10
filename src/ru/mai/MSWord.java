package ru.mai;

/**
 * MS Word using the Apache POI library
 */
public class MSWord {

    /**
     * Start point of program
     */
    public static void main(String[] args) {
        PrimaryMenu primaryMenu = new PrimaryMenu();
        primaryMenu.manageTheMenu("Что вас интересует?\n1. Запись в файл\n2. Чтение из файла\n3. Выход\nВведите цифру: ");
    }
}