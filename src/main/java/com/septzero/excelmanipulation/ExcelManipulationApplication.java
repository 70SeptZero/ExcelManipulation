package com.septzero.excelmanipulation;

import org.apache.commons.compress.utils.Lists;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.util.ObjectUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.apache.poi.ss.usermodel.CellType.FORMULA;

@SpringBootApplication
public class ExcelManipulationApplication {

    public static void main(String[] args) {
        // 获取当前时间
        LocalDateTime currentTime = LocalDateTime.now();
        // 定义日期时间格式
        DateTimeFormatter formatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        // 格式化当前时间为字符串
        String formattedTime = currentTime.format(formatter);
        String outputFilePath = "../output/Manipulation_" + formattedTime + ".xlsx";
        String inputFolderPath = "../excel";

        try {
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet("合并文件");

            File folder = new File(inputFolderPath);
            System.out.println("读取路径为：" + folder.toPath().toAbsolutePath());
            File[] inputFiles = folder.listFiles();

            if (inputFiles != null) {
                for (File inputFile : inputFiles) {
                    System.out.println("fileName:" + inputFile.getName());
                    if (inputFile.isFile()) {
                        Workbook inputWorkbook = WorkbookFactory.create(inputFile);
                        FormulaEvaluator formulaEvaluator = inputWorkbook.getCreationHelper().createFormulaEvaluator();

                        int numberOfSheets = inputWorkbook.getNumberOfSheets();
                        for (int i = 0; i < numberOfSheets; i++) {
                            Sheet inputSheet = inputWorkbook.getSheetAt(i);
                            if (!inputWorkbook.isSheetHidden(i)) {
                                //删除的空行(包含空白行和带有格式的空行)
                                DeleteBlank(inputSheet);
                                int rowCount = inputSheet.getLastRowNum();
                                for (int j = 0; j <= rowCount; j++) {
                                    Row inputRow = inputSheet.getRow(j);
                                    if(inputRow == null){
                                        break;
                                    }
                                    int outRowNum = outputSheet.getLastRowNum()+1;
                                    Row outputRow = outputSheet.createRow(outputSheet.getLastRowNum()+1);

                                    if (inputRow != null) {
                                        int columnCount = inputRow.getLastCellNum();
                                        for (int k = 0; k < columnCount; k++) {
                                            Cell inputCell = inputRow.getCell(k);
                                            Cell outputCell = outputRow.createCell(k);

                                            //把正在处理的单元格信息放进去
                                            Map<String, Object> info = new HashMap<>();
                                            info.put("sheetName", inputSheet.getSheetName());
                                            info.put("fileName", inputFile.getName());
                                            info.put("row", j);
                                            info.put("outRowNum", outRowNum);

                                            if (inputCell != null) {
                                                CellType cellType = inputCell.getCellType();
                                                divCellType(cellType, outputCell, inputCell);
                                                if(cellType == FORMULA){
                                                    getExcelForMula(inputCell, formulaEvaluator, outputCell, info, k);
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        inputWorkbook.close();
                    }
                }
            }

            FileOutputStream outputStream = new FileOutputStream(outputFilePath);
            outputWorkbook.write(outputStream);
            outputWorkbook.close();
            outputStream.close();

            System.out.println("Excel manipulation completed successfully!");
        } catch (IOException e) {
            System.out.println("Error occurred while manipulating excel files: " + e.getMessage());
        }
    }

    private static void getExcelForMula(Cell inputCell, FormulaEvaluator formulaEvaluator, Cell outputCell,
                                        Map<String, Object> info, int k) {
        //处理带公式的单元格
        try{
            //捕获不支持的公式
            CellValue cellValue = formulaEvaluator.evaluate(inputCell);
            if(cellValue != null){
                CellType cellType = cellValue.getCellType();
                divCellType(cellType, outputCell, inputCell);
            }
        }catch (Exception e){
            //遇到不支持的公式，就把公式打印处理啊，然后报出位置
            outputCell.setCellValue(inputCell.getCellFormula());
            String column = numberToLetter(k);
            String row = String.valueOf(info.get("row"));
            String outRowNum = String.valueOf(info.get("outRowNum"));
            String sheetName = String.valueOf(info.get("sheetName"));
            String fileName = String.valueOf(info.get("fileName"));
            System.out.println("——————————!——————————！——————————!——————————");
            System.out.println("!          单元格公式暂不支持，请处理。         !");
            String outInfo = "!文件“" +  fileName + "”的“" + sheetName + "”表，" + column + row + "单元格存在异常!";
            System.out.println(outInfo);
            System.out.println("！       该单元格在输出文件的" + column + outRowNum + "单元格         !");
            System.out.println("——————————!——————————！——————————!——————————");
        }
    }

    private static void DeleteBlank(Sheet inputSheet){
        //去除空白行(带格式的)
        int num = inputSheet.getLastRowNum();
        List<Integer> nums = Lists.newArrayList();
        for (int i = 0; i <= num; i++) {
            Row row = inputSheet.getRow(i);
            boolean flag = true;
            if (row != null) {
                for (Cell cell : row) {
                    //判断该单元格是否为空
                    if (!ObjectUtils.isEmpty(cell.toString())) {
                        flag = false;
                        break;
                    }
                }
                if (flag) {
                    nums.add(i);
                }
                //空白行
            } else {
                nums.add(i);
            }
        }
        //删除无效数据行(带格式的空白行)
        for (Integer n : nums) {
            if (inputSheet.getRow(n) != null) {
                inputSheet.removeRow(inputSheet.getRow(n));
            }
        }
    }

    public static String numberToLetter(int number) {
        //数字转字母，EXCEL定位坐标用的
        // ASCII码中A对应的值
        int asciiA = 65;
        // 拼接，计算字母组合
        StringBuilder sb = new StringBuilder();
        while (number >= 0) {
            int remainder = number % 26; // 获取余数，对应字母的ascii
            char letter = (char) (asciiA + remainder); // 转换为字母
            sb.insert(0, letter); // 将字母插入到字符串的最前面
            number = (number / 26) - 1; // 减去26的倍数，因为从0开始，所以需要减1
            if (number < 0) {
                break; // 如果数字小于0，说明已经完成转换
            }
        }
        return sb.toString();
    }

    public static void divCellType(CellType cellType, Cell outputCell, Cell inputCell){
        //处理不同格式的内容
        switch (cellType) {
            case STRING:
                outputCell.setCellValue(inputCell.getStringCellValue());
                break;
            case NUMERIC:
                outputCell.setCellValue(inputCell.getNumericCellValue());
                break;
            case BOOLEAN:
                outputCell.setCellValue(inputCell.getBooleanCellValue());
                break;
            case FORMULA://这个需要单独处理
            case BLANK:
            case ERROR:
            case _NONE:
                break;
        }
    }

}
