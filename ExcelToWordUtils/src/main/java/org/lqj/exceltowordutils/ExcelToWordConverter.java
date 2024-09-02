package org.lqj.exceltowordutils;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellAlignment;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import java.io.*;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigInteger;
import java.net.URL;
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class ExcelToWordConverter {

    public static void main(String[] args) throws FileNotFoundException {
        URL resource = ExcelToWordConverter.class.getClassLoader().getResource("sheetadmin_002.xlsx");
        String parentPath = new File(resource.getPath()).getParent();
        String inputExcelFile = parentPath+"\\sheetadmin_002.xlsx";
        String inputDocFile = parentPath+"\\test.docx";
        String outputDocFile = parentPath+"\\result.docx";
        appendContentToWord(new FileInputStream(new File(inputExcelFile)),new FileInputStream(new File(inputDocFile)),outputDocFile);
    }

    private static String NUMBERFONTFAMILY = "Times New Roman";
    private static String STRINGFONTFAMILY = "SimSun";
    private static Integer DEFAULTFONTSIZE = 9;

    public static void appendContentToWord(InputStream excelFis, InputStream wordFis,String outputDocFile){
        // 设置允许的最小膨胀比率
        ZipSecureFile.setMinInflateRatio(0.0001);
        try (InputStream fis = excelFis;
             XSSFWorkbook workbook = new XSSFWorkbook(fis);){
            XSSFSheet sheet = workbook.getSheetAt(0); // 假设只处理第一个sheet页
            try (InputStream docFis = wordFis;
                 XWPFDocument doc = new XWPFDocument(docFis)) {
                // 添加一个空行
                doc.createParagraph();
                //查找sheet1中的单元格连续区域
                List<int[]> dataRanges = findContinuousRegions(sheet);
                //查找sheet1中宽度为0的列，用户后续业务跳过单元格
                List<Integer> zeroColIndexs = findWidthZeroColIndexs(sheet);
                if(CollectionUtils.isNotEmpty(dataRanges)){
                    int dataRangesSize = dataRanges.size();
                    for (int i = 0; i < dataRangesSize; i++) {
                        int[] range = dataRanges.get(i);
                        try {
                            // 3. 复制数据和样式到Word文档中
                            copyExcelDataAndStylesToWord(doc, sheet, range, zeroColIndexs,dataRangesSize>1&&i==0);
                        } catch (Exception e) {
                            throw new RuntimeException(e);
                        }
                    }
                    // 添加一个空行
                    doc.createParagraph();
                }
                // 4. 保存Word文档
                saveWordDocument(doc, outputDocFile);
                /*ByteArrayOutputStream out = new ByteArrayOutputStream();
                doc.write(out);
                // 将ByteArrayOutputStream转换为InputStream
                return new ByteArrayInputStream(out.toByteArray());*/
            }
//            System.out.println("转换成功！");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static List<Integer> findWidthZeroColIndexs(XSSFSheet sheet) {
        XSSFRow excelRow = sheet.getRow(0);
        int numCols = excelRow.getPhysicalNumberOfCells();
        List<Integer> zeroColIndexs = new ArrayList<>();
        for (int colIndex = 0; colIndex < numCols; colIndex++) {
            if(sheet.getColumnWidth(colIndex) == 0){
                zeroColIndexs.add(colIndex);
            }
        }
        return zeroColIndexs;
    }
    private static void shiftCellsLeft(Row row, int columnIndexToDelete) {
        int lastColumn = row.getLastCellNum();
        for (int i = columnIndexToDelete; i < lastColumn - 1; i++) {
            Cell cell = row.getCell(i);
            if (cell == null) {
                continue;
            }
            Cell nextCell = row.getCell(i + 1);
            if (nextCell != null) {
                cell.setCellValue(nextCell.getStringCellValue());
            }
        }
        if(lastColumn<1){
            return;
        }
        // 清空最后一个单元格，以便不会复制最后一个单元格的内容到上一个单元格
        Cell lastCell = row.getCell(lastColumn - 1);
        if (lastCell != null) {
            row.removeCell(lastCell);
        }
    }

    public static List<int[]> findContinuousRegions(Sheet sheet) {
        int filterCount=3;
        List<int[]> regions = new ArrayList<>();
        boolean[][] visited = new boolean[sheet.getPhysicalNumberOfRows()][];

        for (int r = 0; r < sheet.getPhysicalNumberOfRows(); r++) {
            Row row = sheet.getRow(r);
            if (row != null && row.getLastCellNum()>0) {
                visited[r] = new boolean[row.getLastCellNum()];
            }
        }

        for (int r = 0; r < sheet.getPhysicalNumberOfRows(); r++) {
            Row row = sheet.getRow(r);
            if (row == null) continue;

            for (int c = 0; c < row.getLastCellNum(); c++) {
                if (visited[r][c]) continue;

                Cell cell = row.getCell(c);
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    visited[r][c] = true;
                    continue;
                }

                int endRow = r;
                int endCol = c;

                for (int rr = r; rr < sheet.getPhysicalNumberOfRows(); rr++) {
                    Row innerRow = sheet.getRow(rr);
                    if (innerRow == null || innerRow.getCell(c) == null || innerRow.getCell(c).getCellType() == CellType.BLANK) {
                        break;
                    }
                    endRow = rr;
                }

                for (int cc = c; cc < row.getLastCellNum(); cc++) {
                    boolean isEmptyColumn = false;
                    for (int rr = r; rr <= endRow; rr++) {
                        Row innerRow = sheet.getRow(rr);
                        if (innerRow == null || innerRow.getCell(cc) == null || innerRow.getCell(cc).getCellType() == CellType.BLANK) {
                            isEmptyColumn = true;
                            break;
                        }
                    }
                    if (isEmptyColumn) {
                        break;
                    }
                    endCol = cc;
                }
                if(endRow-r<filterCount || endCol-c<filterCount){
                    continue;
                }
                regions.add(new int[]{r, c, endRow, endCol});

                for (int rr = r; rr <= endRow; rr++) {
                    for (int cc = c; cc <= endCol; cc++) {
                        visited[rr][cc] = true;
                    }
                }
            }
        }

        return regions;
    }

    /**
     * 从Excel复制数据和样式到Word文档
     */
    private static void copyExcelDataAndStylesToWord(XWPFDocument doc,XSSFSheet sheet,int[] dataRange,List<Integer> zeroColIndexs,  boolean removeBorder) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        int startRow = Integer.valueOf(dataRange[0]);
        int endRow = Integer.valueOf(dataRange[2]);
        int startCol = Integer.valueOf(dataRange[1]);
        int endCol = Integer.valueOf(dataRange[3]);
        XWPFTable table = doc.createTable(endRow-startRow+1,endCol-startCol-zeroColIndexs.size()+1); // 创建新的表格
        Set<String> cellMergedRegions = new HashSet<>();
        for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
            XWPFTableRow wordRow = table.getRow(rowIndex-startRow); // 创建新的表格行
            if (wordRow == null) {
                wordRow = table.createRow();
            }
            XSSFRow excelRow = sheet.getRow(rowIndex);
            if (excelRow.getHeight() != -1) {
                wordRow.setHeight((int) (excelRow.getHeight())); // 将 Excel 高度转换为 Twips 单位
            }
            //整理行中列宽度数组
            int numCols = excelRow.getPhysicalNumberOfCells();
            int[] colWidths = null;
            if(colWidths == null){
                colWidths = new int[numCols];
                for (int colIndex = 0; colIndex < numCols; colIndex++) {
                    colWidths[colIndex] = sheet.getColumnWidth(colIndex);
                }
            }

            int deleteColIndexCount = 0;
            for (int colIndex = startCol; colIndex <= endCol; colIndex++) {
                int wordCellIndex = colIndex-startCol-deleteColIndexCount;
                if(zeroColIndexs.contains(colIndex)){
                    deleteColIndexCount++;
                    continue;
                }
                int colWidth = colWidths[colIndex];
                XWPFTableCell wordCell = wordRow.getCell(wordCellIndex); // 创建新的表格单元格
                if (wordCell == null) {
                    wordCell = wordRow.createCell();
                }
                XSSFCell excelCell = excelRow.getCell(colIndex);
                if (excelCell != null) {
                    // 复制单元格样式
                    copyCellValueAndStyle(excelCell, colWidth, wordCell);
                    // 设置单元格样式
//                    if (excelCell != null && excelCell.getSheet().getNumMergedRegions() > 0) {
//                        CellRangeAddress mergedRegion = getMergedRegion(sheet, rowIndex, colIndex);
//                        if (mergedRegion != null) {
//                            String cellMergedRegion =  generateCellMargin(mergedRegion,zeroColIndexs,startRow,startCol);
//                            cellMergedRegions.add(cellMergedRegion);
//                        }
//                    }
                }
            }
        }
        if(removeBorder){
            removeTableBorders(table);
        }
        List<String> collect = sheet.getMergedRegions().stream().map(m -> generateCellMargin(m, zeroColIndexs, startRow, startCol)).collect(Collectors.toList());
        cellMergedRegions.addAll(collect);
        if(CollectionUtils.isNotEmpty(cellMergedRegions)){
            cellMergedRegions.forEach(cms->{
                String[] splits = cms.split(",");
                mergeCells(table,Integer.valueOf(splits[0]),Integer.valueOf(splits[1]),Integer.valueOf(splits[2]),Integer.valueOf(splits[3]));
            });
        }
    }

    private static String generateCellMargin(CellRangeAddress mergedRegion, List<Integer> zeroColIndexs,int startRow,int startCol) {
        int rowStart = mergedRegion.getFirstRow() - startRow;
        int rowEnd = mergedRegion.getLastRow() - startRow;
        int colStart = mergedRegion.getFirstColumn() - generateMinusCount(mergedRegion.getFirstColumn(),zeroColIndexs) - startCol;
        int colEnd = mergedRegion.getLastColumn() - generateMinusCount(mergedRegion.getLastColumn(),zeroColIndexs) - startCol;
        return String.format("%d,%d,%d,%d", rowStart, colStart, rowEnd, colEnd);
    }

    private static int generateMinusCount(int colIndex, List<Integer> zeroColIndexs) {
        if(CollectionUtils.isEmpty(zeroColIndexs)){
            return 0;
        }
        int count = 0;
        for (Integer zeroColIndex : zeroColIndexs) {
            if(zeroColIndex<colIndex){
                count++;
            }
        }
        return count;
    }

    private static boolean isCellInHiddenMergedRegion(Sheet sheet, Cell cell) {
        int rowIndex = cell.getRowIndex();
        int columnIndex = cell.getColumnIndex();
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        for (CellRangeAddress region : mergedRegions) {
            if (region.isInRange(rowIndex, columnIndex)) {
                // Check if the merged region is hidden (either row or column is hidden)
                for (int i = region.getFirstRow(); i <= region.getLastRow(); i++) {
                    if (sheet.getRow(i).getZeroHeight()) {
                        return true;
                    }
                }
                for (int j = region.getFirstColumn(); j <= region.getLastColumn(); j++) {
                    if (sheet.isColumnHidden(j)) {
                        return true;
                    }
                }
            }
        }
        return false;
    }

    public static void mergeCells(XWPFTable table, int top, int left, int bottom,int right) {
        for(int i= top;i<=bottom;i++){
            mergeCellsHorizontal(table,i,left,right);
        }
        for(int i= left;i<=right;i++){
            mergeCellsVertically(table,i,top,bottom);
        }
    }


    // 水平合并单元格
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        XWPFTableRow tableRow = table.getRow(row);
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            if(tableRow == null){
                continue;
            }
            XWPFTableCell cell = tableRow.getCell(cellIndex);
            if(cell==null){
                continue;
            }
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableRow row = table.getRow(rowIndex);
            if(row == null){
                continue;
            }
            XWPFTableCell cell = row.getCell(col);
            if(cell==null){
                continue;
            }
            if (rowIndex == fromRow) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    // 获取单元格合并区域
    private static CellRangeAddress getMergedRegion(Sheet sheet, int rowIdx, int colIdx) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress merged = sheet.getMergedRegion(i);
            if (merged.isInRange(rowIdx, colIdx)) {
                return merged;
            }
        }
        return null;
    }

    private static Pattern numberPattern = Pattern.compile("^-?\\d+(\\.\\d+)?$");

    /**
     * 获取单元格的内容
     */
    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    DataFormatter formatter = new DataFormatter();
                    return formatter.formatCellValue(cell);
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            default:
                return "";
        }
    }


    private static void setCellWidthForHidden(XWPFTableCell wordCell){
        CTTc ctTc = wordCell.getCTTc();
        CTTcPr tcPr = ctTc.addNewTcPr();
        CTTblWidth tblWidth = tcPr.addNewTcW();
        tblWidth.setW(BigInteger.valueOf(0));

        // 设置单元格边框颜色为白色
        CTTcBorders borders = tcPr.addNewTcBorders();
        borders.addNewTop().setColor("FFFFFF");
        borders.addNewBottom().setColor("FFFFFF");
        borders.addNewLeft().setColor("FFFFFF");
        borders.addNewRight().setColor("FFFFFF");

        // 添加一个空段落并设置其颜色为白色
        for (XWPFParagraph paragraph : wordCell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                run.setColor("FFFFFF");
            }
        }
    }
    //设置单元格宽度
    private static void setCellWidth(XWPFTableCell wordCell,int colWidth){
        CTTcPr tcPr = wordCell.getCTTc().getTcPr();
        if (tcPr == null) {
            tcPr = wordCell.getCTTc().addNewTcPr();
        }
        CTTblWidth tblWidth = tcPr.getTcW();
        if (tblWidth == null) {
            tblWidth = tcPr.addNewTcW();
        }
        tblWidth.setW(BigInteger.valueOf(colWidth)); // 设置单元格宽度为0
    }
    /**
     * 复制单元格样式
     */
    private static void copyCellValueAndStyle(XSSFCell excelCell, int colWidth, XWPFTableCell wordCell) throws NoSuchMethodException, InvocationTargetException, IllegalAccessException {
        XSSFCellStyle sourceStyle = excelCell.getCellStyle();
        XWPFParagraph paragraph = wordCell.getParagraphs().get(0);
        XWPFRun run = paragraph.createRun();

        String cellValue = getCellValue(excelCell);
        cellValue = cellValue.replace("\n", "").replace("\r"," ").replace("\t","").trim();
        // 复制内容到Word文档中
        run.setText(cellValue);

        Matcher matcher = numberPattern.matcher(cellValue);
        boolean isNumber = matcher.matches();

        //设置单元格宽度
        setCellWidth(wordCell,colWidth);

        // 通过反射调用受保护的 getCellAlignment 方法
        Method getCellAlignmentMethod = XSSFCellStyle.class.getDeclaredMethod("getCellAlignment");
        getCellAlignmentMethod.setAccessible(true);
        // 获取水平对齐和垂直对齐方式
        HorizontalAlignment alignment = ((XSSFCellAlignment) getCellAlignmentMethod.invoke(sourceStyle)).getHorizontal();
        if(alignment == HorizontalAlignment.CENTER){
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        } else if (alignment == HorizontalAlignment.CENTER) {
            paragraph.setAlignment(ParagraphAlignment.CENTER);
        }else if(alignment == HorizontalAlignment.LEFT){
            paragraph.setAlignment(ParagraphAlignment.LEFT);
        }

        VerticalAlignment verticalAlignment = sourceStyle.getVerticalAlignment();
        if(verticalAlignment == VerticalAlignment.CENTER || verticalAlignment == VerticalAlignment.BOTTOM){
            CTTcPr ctTcPr = getCTTcPr(wordCell);
            ctTcPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        }

        // 设置字体样式
        XSSFFont font = sourceStyle.getFont();

        CTRPr rpr = run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts = rpr.addNewRFonts();
        if(isNumber){
            run.setFontFamily(NUMBERFONTFAMILY);
            fonts.setEastAsia(NUMBERFONTFAMILY);
            fonts.setAscii(NUMBERFONTFAMILY);
            fonts.setHAnsi(NUMBERFONTFAMILY);
        }else{
            run.setFontFamily(STRINGFONTFAMILY);
            fonts.setEastAsia(STRINGFONTFAMILY);
            fonts.setAscii(STRINGFONTFAMILY);
            fonts.setHAnsi(STRINGFONTFAMILY);
        }
        run.setFontSize(DEFAULTFONTSIZE);
//            run.setFontSize(font.getFontHeightInPoints());
//            run.setFontFamily(font.getFontName());
        if (font != null) {
            run.setBold(font.getBold());
            run.setItalic(font.getItalic());
            if (font.getUnderline() == Font.U_SINGLE) {
                run.setUnderline(UnderlinePatterns.SINGLE);
            }
        }

        // 设置单元格背景色
        if (sourceStyle.getFillPattern() == FillPatternType.SOLID_FOREGROUND) {
            wordCell.setColor(getXWPFColor(sourceStyle.getFillForegroundColorColor()));
        }

        // 设置边框
        setCellBorders(sourceStyle, wordCell);
    }

    /**
     * 获取XWPF颜色对象
     */
    private static String getXWPFColor(Color color) {
        if (color instanceof XSSFColor) {
            byte[] rgb = ((XSSFColor) color).getRGB();
            if (rgb != null) {
                int red = rgb[0] & 0xFF;
                int green = rgb[1] & 0xFF;
                int blue = rgb[2] & 0xFF;
                return String.format("%02X%02X%02X", red, green, blue);
            }
        }
        return null;
    }

    private static CTTcPr getCTTcPr(XWPFTableCell targetCell){
        CTTc ctTc = targetCell.getCTTc();
        CTTcPr ctTcPr = ctTc.isSetTcPr() ? ctTc.getTcPr() : ctTc.addNewTcPr();
        return ctTcPr;
    }
    private static void removeTableBorders(XWPFTable table) {
        // 设置表格的整体边框为无
        CTTblPr tblPr = table.getCTTbl().getTblPr();
        if (tblPr == null) {
            tblPr = table.getCTTbl().addNewTblPr();
        }
        CTTblBorders tblBorders = tblPr.getTblBorders();
        if (tblBorders == null) {
            tblBorders = tblPr.addNewTblBorders();
        }
        tblBorders.addNewTop().setVal(STBorder.NONE);
        tblBorders.addNewBottom().setVal(STBorder.NONE);
        tblBorders.addNewLeft().setVal(STBorder.NONE);
        tblBorders.addNewRight().setVal(STBorder.NONE);
        tblBorders.addNewInsideH().setVal(STBorder.NONE);
        tblBorders.addNewInsideV().setVal(STBorder.NONE);

        // 设置每个单元格的边框为无
        for (XWPFTableRow row : table.getRows()) {
            for (XWPFTableCell cell : row.getTableCells()) {
                CTTcPr tcPr = cell.getCTTc().getTcPr();
                if (tcPr == null) {
                    tcPr = cell.getCTTc().addNewTcPr();
                }
                CTTcBorders borders = tcPr.isSetTcBorders() ? tcPr.getTcBorders() : tcPr.addNewTcBorders();
                borders.addNewTop().setVal(STBorder.NONE);
                borders.addNewBottom().setVal(STBorder.NONE);
                borders.addNewLeft().setVal(STBorder.NONE);
                borders.addNewRight().setVal(STBorder.NONE);
                borders.addNewInsideH().setVal(STBorder.NONE);
                borders.addNewInsideV().setVal(STBorder.NONE);
            }
        }
    }
    /**
     * 设置单元格边框
     */
    private static void setCellBorders(CellStyle sourceStyle, XWPFTableCell targetCell) {
        CTTcPr ctTcPr = getCTTcPr(targetCell);
        // 下边框
        if (sourceStyle.getBorderBottom() == BorderStyle.NONE) {
            CTTcBorders ctTcBorders = ctTcPr.isSetTcBorders() ? ctTcPr.getTcBorders() : ctTcPr.addNewTcBorders();
            CTBorder ctBorder = ctTcBorders.isSetBottom() ? ctTcBorders.getBottom() : ctTcBorders.addNewBottom();
            ctBorder.setVal(STBorder.NONE);
        }

        // 上边框
        if (sourceStyle.getBorderTop() == BorderStyle.NONE) {
            CTTcBorders ctTcBorders = ctTcPr.isSetTcBorders() ? ctTcPr.getTcBorders() : ctTcPr.addNewTcBorders();
            CTBorder ctBorder = ctTcBorders.isSetTop() ? ctTcBorders.getTop() : ctTcBorders.addNewTop();
            ctBorder.setVal(STBorder.NONE);
        }

        // 左边框
        if (sourceStyle.getBorderLeft() == BorderStyle.NONE) {
            CTTcBorders ctTcBorders = ctTcPr.isSetTcBorders() ? ctTcPr.getTcBorders() : ctTcPr.addNewTcBorders();
            CTBorder ctBorder = ctTcBorders.isSetLeft() ? ctTcBorders.getLeft() : ctTcBorders.addNewLeft();
            ctBorder.setVal(STBorder.NONE);
        }

        // 右边框
        if (sourceStyle.getBorderRight() == BorderStyle.NONE) {
            CTTcBorders ctTcBorders = ctTcPr.isSetTcBorders() ? ctTcPr.getTcBorders() : ctTcPr.addNewTcBorders();
            CTBorder ctBorder = ctTcBorders.isSetRight() ? ctTcBorders.getRight() : ctTcBorders.addNewRight();
            ctBorder.setVal(STBorder.NONE);
        }
    }

    /**
     * 保存Word文档
     */
    private static void saveWordDocument(XWPFDocument doc, String outputFile) throws IOException {
        FileOutputStream out = null;
        try {
            File file = new File(outputFile);
            if(!file.exists()){
                file.createNewFile();
            }
            out = new FileOutputStream(file);
            doc.write(out);
        } finally {
            if (out != null) {
                out.close();
            }
            doc.close();
        }
    }
}
