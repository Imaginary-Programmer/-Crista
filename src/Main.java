import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {
    public static void main(String[] args) throws Exception {
        ArrayList<String> valuesColumn = new ArrayList<>();
        int[][] data = readExcel(valuesColumn);
        if (data != null) {
            Scanner scan = new Scanner(System.in);
            System.out.print("Enter how you want to display the result:\n1 - to excel file\n2 - to console" +
                    "\nYour choice: ");
            String resultOutput = scan.nextLine();
            switch (resultOutput) {
                case "1":
                    System.out.print("Enter the path for the new excel file: ");
                    String path = scan.nextLine();
                    Workbook workbookWrite = new XSSFWorkbook();
                    Sheet newSheet = workbookWrite.createSheet("Result");
                    dataProcessing(data, data.length, valuesColumn.size(), valuesColumn, path, workbookWrite,
                            newSheet, resultOutput);
                    System.out.println("File created.");
                    break;
                case "2":
                    dataProcessing(data, data.length, valuesColumn.size(), valuesColumn, null, null,
                            null, resultOutput);
                    break;
                default:
                    System.out.println("There is no such choice!!!");
                    break;
            }
        }
        valuesColumn.clear();
    }

    public static int[][] readExcel(ArrayList<String> valuesColumn)///считывае данные из excel файла: критерии и чисел
    {
        try {
            Scanner scan = new Scanner(System.in);
            System.out.print("Enter the path to the excel file: ");
            String path = scan.nextLine();
            FileInputStream file = new FileInputStream(new File(path));
            XSSFWorkbook workbookRead = new XSSFWorkbook(file);
            System.out.print("Enter page number: ");
            int sheetNumber = scan.nextInt();
            XSSFSheet sheet = workbookRead.getSheetAt(sheetNumber);
            Iterator<Row> rowIterator = sheet.iterator();
            //////считывание данных//////
            int[][] numbers = new int[(sheet.getLastRowNum())][(sheet.getRow(0).getPhysicalNumberOfCells())];
            while (rowIterator.hasNext()) {///заполнять массив до тех пор, пока строка в excel файле заполнена какими-либо данными
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {
                        case NUMERIC:
                            if (row.getRowNum() > 0) {
                                numbers[(row.getRowNum() - 1)][cell.getColumnIndex()] = (int) cell.getNumericCellValue();
                            }
                            break;
                        case STRING:
                            if (row.getRowNum() == 0) {
                                valuesColumn.add(cell.getStringCellValue());
                            }
                            break;
                    }
                }
            }
            return numbers;
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
            return null;
        }
    }

    public static void dataProcessing(int[][] data, int row, int column, ArrayList<String> valuesColumn, String path,
                                      Workbook workbookWrite, Sheet newSheet, String resultOutput)///обработка критериев и массива с числами
    {
        ArrayList<Integer> indexNum = new ArrayList<>();///индексы колонок с критериями, по которым будут формироваться группы
        ArrayList<String> result = new ArrayList<>();///"строка" данных для записи в excel файл или в консоль
        for (int i = 0; i < valuesColumn.size(); i++) {
            if (!valuesColumn.get(i).equals("-")) {
                result.add(valuesColumn.get(i));
            }
            if (valuesColumn.get(i).equals("NUM")) {
                indexNum.add(i);
            }
        }
        int resultRow = 0;///"указатель" на строку в новом excel файле
        if (resultOutput.equals("1")) {///вывод критериев выбранным способом (на консоль или в файл)
            printExcel(result, resultRow, path, workbookWrite, newSheet);
        } else {
            System.out.println("Result:");
            printConsole(result);
        }
        resultRow++;///смещаем "указатель" в новом excel файле на следующую строку
        result.clear();
        Integer checkedRows = 0;///текущая проверяемая строка
        ArrayList<Integer> curCheckNumber = new ArrayList<>();///индексы группы
        HashSet<Integer> checkRowAll = new HashSet<>();///множество для хранения индексов строк, который уже входят в какую-либо группу
        ArrayList<Integer> curIndexRowElementsGroup = new ArrayList<>();///индексы строк текущей группы
        while (!checkedRows.equals(row))///выполнять пока не будут проверены все строки в матрице
        {
            if (!checkRowAll.contains(checkedRows)) {///выполнять, если текущая строка раннее не входила в какую-ибо группу
                for (int j = 0; j < indexNum.size(); j++) {///записываем критерии в временный список
                    curCheckNumber.add(data[checkedRows][indexNum.get(j)]);
                }
                checkRowAll.add(checkedRows);///добавляем в множество индекс текущей строки
                curIndexRowElementsGroup.add(checkedRows);///добавляем индекс текущей строки в группу
                boolean match = false;
                for (int curRow = checkedRows + 1; curRow < row; curRow++) {///находим строки и запомнимаем индекс,
                    // которые подходят под все критерии
                    match = true;
                    for (int j = 0; j < indexNum.size(); j++) {
                        if (!curCheckNumber.get(j).equals(data[curRow][indexNum.get(j)]) || (checkRowAll.contains(curRow))) {
                            match = false;
                        }
                    }
                    if (match) {
                        curIndexRowElementsGroup.add(curRow);
                        checkRowAll.add(curRow);
                    }
                }////////////конец заполнения индексов элементов группы
                result = fillRowDataResult(data, valuesColumn, checkedRows, curIndexRowElementsGroup);///заполняем данные об текущей группе
                if (resultOutput.equals("1"))///вывод группы выбранным способом (на консоль или в файл)
                {
                    printExcel(result, resultRow, path, workbookWrite, newSheet);
                } else {
                    printConsole(result);
                }
                resultRow++;///смещаем "указатель" в новом excel файле на следующую строку
                result.clear();///очищаем "строку" с данными о группе
                curCheckNumber.clear();///очищаем индексы группы
                curIndexRowElementsGroup.clear();
            }
            checkedRows++;///смещаем "указатель" в матрице на следующую строку
        }
    }

    public static ArrayList<String> fillRowDataResult(int[][] data, ArrayList<String> valuesColumn,
                                                      int checkedRows, ArrayList<Integer> curIndexRowElementsGroup)///заполняем данные об текущей группе
    {
        ArrayList<String> curResult = new ArrayList<>();
        int curColumnResult = 0;
        int sum = 0;
        int max = 0;
        int min = 0;
        String konk = "";
        while (curColumnResult < valuesColumn.size())///добавляем в временный список все колонки
        {
            if (!valuesColumn.get(curColumnResult).equals("-"))///выполнять, если эта колонка не с критерием "-"
            {
                switch (valuesColumn.get(curColumnResult)) {
                    case "NUM":///просто записываем 1ую строчку данного столбца (они все одинаковы) группы
                        curResult.add(Integer.toString(data[checkedRows][curColumnResult]));
                        break;
                    case "SUM":///суммируем все строчки группы данного столбца
                        sum = 0;
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            sum += (int) data[curIndexRowElementsGroup.get(j)][curColumnResult];
                        }
                        curResult.add(Integer.toString(sum));
                        break;
                    case "MIN":///находим минимальное значение среди всех строчек группы данного столбца
                        min = data[checkedRows][curColumnResult];
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            if (min > data[curIndexRowElementsGroup.get(j)][curColumnResult]) {
                                min = data[curIndexRowElementsGroup.get(j)][curColumnResult];
                            }
                        }
                        curResult.add(Integer.toString(min));
                        break;
                    case "MAX":///находим максимальное значение среди всех строчек группы данного столбца
                        max = data[checkedRows][curColumnResult];
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            if (max < data[curIndexRowElementsGroup.get(j)][curColumnResult]) {
                                max = data[curIndexRowElementsGroup.get(j)][curColumnResult];
                            }
                        }
                        curResult.add(Integer.toString(max));
                        break;
                    case "CONC":///производим конкатенацию всех строчек группы данного столбца
                        konk = "";
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            konk += Integer.toString(data[curIndexRowElementsGroup.get(j)][curColumnResult]);
                        }
                        curResult.add(konk);
                        break;
                }
            }
            curColumnResult++;
        }
        return curResult;
    }

    public static void printExcel(ArrayList<String> data, int indexRow, String path,
                                  Workbook workbookWrite, Sheet newSheet)///запись полученного списка в новый excel файл
    {
        Row row;
        row = newSheet.createRow(indexRow);
        for (int i = 0; i < data.size(); i++) {
            row.createCell(i).setCellValue(data.get(i));
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(path);
            workbookWrite.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }

    public static void printConsole(ArrayList<String> data)///вывод полученного списка на консоль
    {
        for (int i = 0; i < data.size(); i++) {
            System.out.printf("%-10s", data.get(i));
        }
        System.out.println();
    }
}