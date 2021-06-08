package ru.mai;

import java.util.Scanner;

/**
 * Primary Menu
 */
public class Menu {

     /**
      * Scanner object
      */
     private final Scanner num = new Scanner(System.in);

     /**
      * Opening the menu
      */
     public void startMenu() {
          menu("Что вас интересует?\n1. Запись в файл\n2. Чтение из файла\n3. Выход\nВведите цифру: ");
     }

     /**
      * Recursive menu method
      * @param str text condition
      */
     public void menu(String str) {
          int number = 0;
          while (number != 3) {
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
                              menu("Что вас интересует?\n1. Перейти к ручной записи\n2. Перейти к шаблону\n3. Назад\nВведите цифру: ");
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
                    case 3:
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
               if (number <= 0 || number > 3) {
                    System.out.println("Invalid Data! Вы ввели неверное значение!");
                    System.out.print(str);
               }
          } while (number <= 0 || number > 3);
          return number;
     }
}