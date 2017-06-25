package merge4chong;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Forrest on 2017/6/24.
 */
public class MergeExcel {
    public static void mergeXSSFWorkbooks(JTextArea area, File dest, List<File>source) {
        if(source == null || source.size()== 0){
            return;
        }
        try {

            XSSFWorkbook destExcel = new XSSFWorkbook();
            for (File file : source) {
                XSSFWorkbook sourcebook = new XSSFWorkbook(new FileInputStream(file));
                int sheetNums = sourcebook .getNumberOfSheets();
                for (int index = 0; index < sheetNums; index++) {
                    XSSFSheet sourcesheet = sourcebook.getSheetAt(index);
                    XSSFSheet destSheet = destExcel.getSheet(sourcesheet.getSheetName());
                    int firstRowNum = sourcesheet.getFirstRowNum();
                    if (destSheet == null) {
                        destSheet = destExcel.createSheet(sourcesheet.getSheetName());
                    } else {
                        firstRowNum = firstRowNum + 3;//because my friend's excels head line take 3 rows
                    }
                    System.out.println(sourcesheet.getLastRowNum());
                    copyRows(destExcel, sourcebook, sourcesheet, destSheet, firstRowNum, sourcesheet.getLastRowNum(), destSheet.getLastRowNum());
                }
                area.append("已合并文件："+file.getName()+ "\r\n");
            }
            FileOutputStream out = new FileOutputStream(new File(dest.getParentFile().getAbsolutePath() + File.separator + "test.xls"));
            destExcel.write(out);
            out.flush();
            out.close();
            area.append("完成！！\r\n");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * @param destWorkBook 目标workbook
     * @param sourceWorkBook 源workbook
     * @param sourceSheet 源sheet
     * @param targetSheet 目sheet
     * @param pStartRow 起始读取行
     * @param pEndRow 结束读取行
     * @param pPosition 目标保存
     */
    public static void copyRows(XSSFWorkbook destWorkBook, XSSFWorkbook sourceWorkBook, XSSFSheet sourceSheet, XSSFSheet targetSheet, int pStartRow, int pEndRow, int pPosition) {
        XSSFRow sourceRow = null;
        XSSFRow targetRow = null;
        XSSFCell sourceCell = null;
        XSSFCell targetCell = null;
        int cType;
        int i;
        int j;
        int targetRowFrom;
        int targetRowTo;

        if ((pStartRow == -1) || (pEndRow == -1)) {
            return;
        }

        List<CellRangeAddress> oldRanges = new ArrayList<CellRangeAddress>();
        for (i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
            oldRanges.add(sourceSheet.getMergedRegion(i));
        }

        // 拷贝合并的单元格。原理：复制当前合并单元格后，原位置的格式会移动到新位置，需在原位置生成旧格式
        for (int k = 0; k < oldRanges.size(); k++) {
            CellRangeAddress oldRange = oldRanges.get(k);
            CellRangeAddress newRange = new CellRangeAddress(oldRange
                    .getFirstRow(), oldRange.getLastRow(), oldRange
                    .getFirstColumn(), oldRange.getLastColumn());

            if (oldRange.getFirstRow() >= pStartRow
                    && oldRange.getLastRow() <= pEndRow) {
                targetRowFrom = oldRange.getFirstRow() - pStartRow + pPosition;
                targetRowTo = oldRange.getLastRow() - pStartRow + pPosition;
                oldRange.setFirstRow(targetRowFrom);
                oldRange.setLastRow(targetRowTo);
                targetSheet.addMergedRegion(oldRange);
                sourceSheet.addMergedRegion(newRange);
            }
        }
        // 设置列宽
        for (i = pStartRow; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow != null) {
                for (j = sourceRow.getLastCellNum(); j > sourceRow.getFirstCellNum(); j--) {
                    targetSheet.setColumnWidth(j, sourceSheet.getColumnWidth(j));
                    targetSheet.setColumnHidden(j, false);
                }
                break;
            }
        }
        // 拷贝行并填充数据
        for (; i <= pEndRow; i++) {
            sourceRow = sourceSheet.getRow(i);
            if (sourceRow == null) {
                continue;
            }
            targetRow = targetSheet.createRow(i - pStartRow + pPosition);
            targetRow.setHeight(sourceRow.getHeight());
            for (j = sourceRow.getFirstCellNum(); j <= sourceRow.getPhysicalNumberOfCells(); j++) {
                sourceCell = sourceRow.getCell(j);
                if (sourceCell == null) {
                    continue;
                }
                targetCell = targetRow.createCell(j);

                //样式的设置

                XSSFCellStyle cStyle = destWorkBook.createCellStyle();
                cStyle.cloneStyleFrom(sourceCell.getCellStyle());
                targetCell.setCellStyle(cStyle);

                cType = sourceCell.getCellType();
                targetCell.setCellType(cType);
                switch (cType) {
                    case XSSFCell.CELL_TYPE_BOOLEAN:
                        targetCell.setCellValue(sourceCell.getBooleanCellValue());
                        // System.out.println("--------TYPE_BOOLEAN:" + targetCell.getBooleanCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_ERROR:
                        targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
                        // System.out.println("--------TYPE_ERROR:" + targetCell.getErrorCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_FORMULA:
                        // parseFormula这个函数的用途在后面说明
                        targetCell.setCellFormula(parseFormula(sourceCell.getCellFormula()));
                        // System.out.println("--------TYPE_FORMULA:" + targetCell.getCellFormula());
                        break;
                    case XSSFCell.CELL_TYPE_NUMERIC:
                        targetCell.setCellValue(sourceCell.getNumericCellValue());
                        // System.out.println("--------TYPE_NUMERIC:" + targetCell.getNumericCellValue());
                        break;
                    case XSSFCell.CELL_TYPE_STRING:
                        targetCell.setCellValue(sourceCell.getRichStringCellValue());
                        // System.out.println("--------TYPE_STRING:" + i + targetCell.getRichStringCellValue());
                        break;
                }
            }
        }
    }

    /**
     * 处理公式
     * @param pPOIFormula
     * @return
     */
    private static String parseFormula(String pPOIFormula) {
        final String cstReplaceString = "ATTR(semiVolatile)"; //$NON-NLS-1$
        StringBuffer result = null;
        int index;
        result = new StringBuffer();
        index = pPOIFormula.indexOf(cstReplaceString);
        if (index >= 0) {
            result.append(pPOIFormula.substring(0, index));
            result.append(pPOIFormula.substring(index + cstReplaceString.length()));
        } else {
            result.append(pPOIFormula);
        }
        return result.toString();
    }
}
