package excel;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excel {
    private ArrayList<String> valuesColumn;
    private int[][] data;
    private String pathNew;
    private String resultOutput;
    private Workbook workbookWrite;
    private Sheet newSheet;
    private String path;
    private ArrayList<String> result;
    private int resultRow;

    public Excel() {
        try {
            this.valuesColumn = new ArrayList<>();
            this.data = readExcel();
            if (this.data != null) {
                Scanner scan = new Scanner(System.in);
                System.out.print("Enter how you want to display the result:\n1 - to excel file\n2 - to console" +
                        "\nYour choice: ");
                this.resultOutput = scan.nextLine();
                switch (this.resultOutput) {
                    case "1":
                        System.out.print("Enter the path for the new excel file: ");
                        this.pathNew = scan.nextLine();
                        this.workbookWrite = new XSSFWorkbook();
                        this.newSheet = workbookWrite.createSheet("Result");
                        dataProcessing(data.length, valuesColumn.size());
                        System.out.println("File created.");
                        break;
                    case "2":
                        dataProcessing(this.data.length, this.valuesColumn.size());
                        break;
                    default:
                        System.out.println("There is no such choice!!!");
                        break;
                }
            }
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
        }
    }


    public int[][] readExcel()///считывае данные из excel файла: критерии и чисел
    {
        try {
            Scanner scan = new Scanner(System.in);
            System.out.print("Enter the path to the excel file: ");
            this.path = scan.nextLine();
            FileInputStream file = new FileInputStream(new File(this.path));
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
                                this.valuesColumn.add(cell.getStringCellValue());
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

    public void dataProcessing(int row, int column)///обработка критериев и массива с числами
    {
        ArrayList<Integer> indexNum = new ArrayList<>();///индексы колонок с критериями, по которым будут формироваться группы
        result = new ArrayList<>();///"строка" данных для записи в excel файл или в консоль
        for (int i = 0; i < this.valuesColumn.size(); i++) {
            if (!this.valuesColumn.get(i).equals("-")) {
                this.result.add(this.valuesColumn.get(i));
            }
            if (this.valuesColumn.get(i).equals("NUM")) {
                indexNum.add(i);
            }
        }
        this.resultRow = 0;///"указатель" на строку в новом excel файле
        print();
        this.resultRow++;///смещаем "указатель" в новом excel файле на следующую строку
        this.result.clear();
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
                        if (!curCheckNumber.get(j).equals(this.data[curRow][indexNum.get(j)]) || (checkRowAll.contains(curRow))) {
                            match = false;
                        }
                    }
                    if (match) {
                        curIndexRowElementsGroup.add(curRow);
                        checkRowAll.add(curRow);
                    }
                }////////////конец заполнения индексов элементов группы
                this.result = fillRowDataResult(checkedRows, curIndexRowElementsGroup);///заполняем данные об текущей группе
                print();
                this.resultRow++;///смещаем "указатель" в новом excel файле на следующую строку
                this.result.clear();///очищаем "строку" с данными о группе
                curCheckNumber.clear();///очищаем индексы группы
                curIndexRowElementsGroup.clear();
            }
            checkedRows++;///смещаем "указатель" в матрице на следующую строку
        }
    }

    public ArrayList<String> fillRowDataResult(int checkedRows, ArrayList<Integer> curIndexRowElementsGroup)///заполняем данные об текущей группе
    {
        ArrayList<String> curResult = new ArrayList<>();
        int curColumnResult = 0;
        int sum = 0;
        int max = 0;
        int min = 0;
        String konk = "";
        while (curColumnResult < this.valuesColumn.size())///добавляем в временный список все колонки
        {
            if (!this.valuesColumn.get(curColumnResult).equals("-"))///выполнять, если эта колонка не с критерием "-"
            {
                switch (this.valuesColumn.get(curColumnResult)) {
                    case "NUM":///просто записываем 1ую строчку данного столбца (они все одинаковы) группы
                        curResult.add(Integer.toString(this.data[checkedRows][curColumnResult]));
                        break;
                    case "SUM":///суммируем все строчки группы данного столбца
                        sum = 0;
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            sum += (int) this.data[curIndexRowElementsGroup.get(j)][curColumnResult];
                        }
                        curResult.add(Integer.toString(sum));
                        break;
                    case "MIN":///находим минимальное значение среди всех строчек группы данного столбца
                        min = this.data[checkedRows][curColumnResult];
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            if (min > this.data[curIndexRowElementsGroup.get(j)][curColumnResult]) {
                                min = this.data[curIndexRowElementsGroup.get(j)][curColumnResult];
                            }
                        }
                        curResult.add(Integer.toString(min));
                        break;
                    case "MAX":///находим максимальное значение среди всех строчек группы данного столбца
                        max = this.data[checkedRows][curColumnResult];
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            if (max < this.data[curIndexRowElementsGroup.get(j)][curColumnResult]) {
                                max = this.data[curIndexRowElementsGroup.get(j)][curColumnResult];
                            }
                        }
                        curResult.add(Integer.toString(max));
                        break;
                    case "CONC":///производим конкатенацию всех строчек группы данного столбца
                        konk = "";
                        for (int j = 0; j < curIndexRowElementsGroup.size(); j++) {
                            konk += Integer.toString(this.data[curIndexRowElementsGroup.get(j)][curColumnResult]);
                        }
                        curResult.add(konk);
                        break;
                }
            }
            curColumnResult++;
        }
        return curResult;
    }

    public void print()///вывод данных выбранным способом (на консоль или в файл)
    {
        if (this.resultOutput.equals("1")) {
            printExcel(this.result, this.resultRow);
        } else {
            if (this.resultRow == 0) {
                System.out.println("Result:");
            }
            printConsole(this.result);
        }
    }

    public void printExcel(ArrayList<String> data, int indexRow)///запись полученного списка в новый excel файл
    {
        Row row;
        row = this.newSheet.createRow(indexRow);
        for (int i = 0; i < data.size(); i++) {
            row.createCell(i).setCellValue(data.get(i));
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(this.pathNew);
            this.workbookWrite.write(fileOut);
            fileOut.close();
        } catch (Exception e) {
            System.out.println("Error: " + e.getMessage());
            System.exit(1);
        }
    }

    public void printConsole(ArrayList<String> data)///вывод полученного списка на консоль
    {
        for (int i = 0; i < data.size(); i++) {
            System.out.printf("%-10s", data.get(i));
        }
        System.out.println();
    }
}