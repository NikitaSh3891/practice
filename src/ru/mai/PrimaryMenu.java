package ru.mai;

import java.util.Scanner;

/**
 * Primary menu for working with files
 */
public class PrimaryMenu {

     /**
      * Scanner object, allows you to read data from the console
      */
     private final Scanner num = new Scanner(System.in);

     /**
      * Number of menu items
      */
     private final int NUMBER_OF_MENU = 3;

     /**
      * Recursive file type selection menu
      * @param str text condition
      */
     public void manageTheMenu(String str) {
          int number = 0;
          while (number != NUMBER_OF_MENU) {
               if (number == 0) {
                    System.out.print(str);
               }
               number = inputCheck(str);
               switch (number) {
                    case 1:
                         if (str.equals("Что вас интересует?\n1. Перейти к ручной записи\n2. Перейти к шаблону\n3. Назад\nВведите цифру: ")) {
                              InputWord inputWord = new InputWord();
                              inputWord.menu();
                         } else {
                              manageTheMenu("Что вас интересует?\n1. Перейти к ручной записи\n2. Перейти к шаблону\n3. Назад\nВведите цифру: ");
                         }
                         number = 0;
                         break;
                    case 2:
                         if (str.equals("Что вас интересует?\n1. Запись в файл\n2. Чтение из файла\n3. Выход\nВведите цифру: ")){
                              OutputWord outputWord = new OutputWord();
                              outputWord.outputWord();
                         }
                         if (str.equals("Что вас интересует?\n1. Перейти к ручной записи\n2. Перейти к шаблону\n3. Назад\nВведите цифру: ")) {
                              InputWordTemplate inputWordTemplate = new InputWordTemplate();
                              inputWordTemplate.inputWordTemplate();
                         }
                         number = 0;
                         break;
                    case NUMBER_OF_MENU:
                         break;
                    default:
                         System.out.println("Invalid Data! Неверное значение!");
               }
          }
     }

     /**
      * Checking the selection of menu items
      * @param str text condition
      * @return number menu item
      */
     private int inputCheck(String str) {
          int number;
          do {
               while (!num.hasNextInt()) {
                    System.out.println("Invalid Data! Вы ввели неверное значение!");
                    num.next();
                    System.out.print(str);
               }
               number = num.nextInt();
               if (number <= 0 || number > NUMBER_OF_MENU) {
                    System.out.println("Invalid Data! Вы ввели неверное значение!");
                    System.out.print(str);
               }
          } while (number <= 0 || number > NUMBER_OF_MENU);
          return number;
     }
}