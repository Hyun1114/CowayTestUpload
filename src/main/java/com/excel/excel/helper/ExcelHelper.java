package com.excel.excel.helper;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import com.excel.excel.model.Coway;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

public class ExcelHelper {
    public static String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    static String[] HEADERs = { "Number", "TestContent", "TestType", "GroupColor", "ConUse", "BlockName", "TestApi", "TestId", "TestTime", 
    "EcMin", "EcMax", "VcMin", "VcMax", "NgMin", "NgMax", "VaMin", "VaMax", "VaFix", "TestInfo"};
    static String SHEET = "Coway2";
    public static boolean hasExcelFormat(MultipartFile file) {
        if (!TYPE.equals(file.getContentType())) {
            return false;
        }       
        return true;
    }

    public static List<Coway> excelToCoways(InputStream is) {
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(is);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rows = sheet.iterator();
            
            List<Coway> coways = new ArrayList<Coway>();
            int rowNumber = 0;
            while (rows.hasNext()) {
                Row currentRow = rows.next();

                // skip header
                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }
                Iterator<Cell> cellsInRow = currentRow.iterator();
                Coway coway = new Coway();
                int cellIdx = 0;
                while (cellsInRow.hasNext()) {
                    Cell currentCell = cellsInRow.next();

                    switch (cellIdx) {
                    case 0:
                        coway.setNumber((long) currentCell.getNumericCellValue());
                          break;

                    case 1:
                        coway.setTestContent(currentCell.getStringCellValue());
                        break;

                    case 2:
                        coway.setTestType(currentCell.getStringCellValue());
                        break;

                    case 3:
                        coway.setGroupColor(currentCell.getStringCellValue());
                        break;

                    case 4:
                        coway.setConUse(currentCell.getStringCellValue());
                        break;

                    case 5:
                        coway.setBlockName(currentCell.getStringCellValue());
                        break;

                    case 6:
                        coway.setTestApi(currentCell.getStringCellValue());
                        break;

                    case 7:
                        coway.setTestId(currentCell.getStringCellValue());
                        break;

                    case 8:
                        coway.setTestTime(currentCell.getStringCellValue());
                        break;

                    case 9:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setEcMin(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setEcMin(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 10:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setEcMax(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setEcMax(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 11:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setVcMin(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setVcMin(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;
                            
                    case 12:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setVcMax(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setVcMax(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 13:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setNgMin(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setNgMin(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 14:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setNgMax(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setNgMax(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 15:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setVaMin(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setVaMin(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 16:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setVaMax(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setVaMax(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 17:
                        if(currentCell.getCellType() == CellType.STRING) {
                            coway.setVafix(currentCell.getStringCellValue());
                        } else if(currentCell.getCellType() == CellType.NUMERIC) {
                            coway.setVafix(String.valueOf(currentCell.getNumericCellValue()));
                        }
                        break;

                    case 18:
                        coway.setTestInfo(currentCell.getStringCellValue());
                        break;
                        default:
                        break;
                      }

                    cellIdx++;
                }
                coways.add(coway);
                System.out.println("cellIdx" + cellIdx);
                System.out.println("tutorial" + coway);
          }
            workbook.close();
            return  coways;
            } catch (IOException e) {
            throw new RuntimeException("fail to parse Excel file: " + e.getMessage());
        }
    }
}