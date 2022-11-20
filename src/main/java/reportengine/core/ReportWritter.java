package reportengine.core;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import reportengine.annotations.ColumnReport;
import reportengine.annotations.Report;
import reportengine.enums.ColorStyleFor;

public class ReportWritter<T> {

    private static final Logger LOGGER = LoggerFactory.getLogger(ReportWritter.class);
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final XSSFCellStyle cellStyle = workbook.createCellStyle();

    private XSSFColor colorBorder;
    private XSSFColor colorCell;
    private XSSFFont fontStyle;

    public ReportWritter() {
        // Constructor default
    }

    public void writeInXlsx(T obj, List<T> data) throws IOException {
        String sheetName = obj.getClass().getAnnotation(Report.class).sheetName();
        XSSFSheet sheet = workbook.createSheet(sheetName);
        List<List<Object>> registers = new ArrayList<>();
        Field[] fields = obj.getClass().getDeclaredFields();

        buildContent(data, registers, fields);
        putData(registers, fields, sheet, obj);
    }

    private void buildContent(List<T> objs, List<List<Object>> registers, Field[] fields) {
        List<Object> attributeFields = new ArrayList<>();
        List<Object> attributeValues = new ArrayList<>();

        makeTitle(attributeFields, fields);
        makeDataValues(objs, attributeValues, fields);

        registers.add(attributeFields);
        registers.add(attributeValues);

    }

    private void makeTitle(List<Object> attributeFields, Field[] fields) {
        for(Field field: fields) {
            String fieldName = field.getName().substring(0, 1).toUpperCase() + field.getName().substring(1);
            if(field.isAnnotationPresent(ColumnReport.class)) {
                fieldName = field.getAnnotation(ColumnReport.class).title();
            }
            attributeFields.add(fieldName);
        }
    }

    private void makeDataValues(List<T> objs, List<Object> attributeValues, Field[] fields) {
        objs.forEach(obj -> {
            for(Field field: fields) {
                field.setAccessible(true);
                try {
                    if (obj != null) {
                        attributeValues.add(field.get(obj) != null ? field.get(obj) : " - ");
                    } else {
                        attributeValues.add("-");
                    }
                } catch (IllegalAccessException e) {
                    LOGGER.error(e.getMessage());
                }
            }
        });
    }

    private void putData(List<List<Object>> registers, Field[] fields, XSSFSheet sheet, T obj) throws IOException {
        int rowCount = 0;
        int rowTitle = 0;

        String name = obj.getClass().getAnnotation(Report.class).name();

        for(List<Object> registerObjects: registers) {
            XSSFRow row = sheet.createRow(rowCount++);
            int cellCount = 0;

            for(Object value: registerObjects) {

                if(cellCount >= fields.length) {
                    cellCount = 0;
                    row = sheet.createRow(rowCount++);
                }

                XSSFCell cell = row.createCell(cellCount++);
                cell.setCellValue(value.toString());

                if(rowCount == 1) {
                    applyTitleStyle(sheet, rowTitle, cellCount);
                    applyFilter(sheet, rowTitle, rowTitle, registerObjects.size() -1 );
                }
            }

        }

        File pathXlsx = new File("src/main/resources/static/"+name+".xlsx");
        FileOutputStream outPut = new FileOutputStream(pathXlsx);
        workbook.write(outPut);
        outPut.close();
    }

    private void applyTitleStyle(XSSFSheet sheet, int lineIndex, int cellIndex) {
        sheet.setDefaultColumnWidth(25);
        sheet.setAutobreaks(true);

        XSSFCell cell = sheet.getRow(lineIndex).getCell(--cellIndex);
        cell.setCellStyle(cellStyle);
    }


    public void applyFont(double fontHeight, boolean isBold, XSSFColor fontColor, String fontName, boolean isItalic) {
        XSSFFont font = workbook.createFont();
        font.setFontHeight(fontHeight);
        font.setBold(isBold);
        font.setColor(fontColor);
        font.setFontName(fontName);
        font.setItalic(isItalic);
        this.cellStyle.setFont(font);
    }

    public void applyCellStyle(XSSFColor colorCell, FillPatternType patternType, HorizontalAlignment alignment) {
        this.cellStyle.setFillForegroundColor(colorCell);
        this.cellStyle.setFillPattern(patternType);
        this.cellStyle.setAlignment(alignment);
    }

    /**
     * Apply border for selected cell
     * @param borderStyle type border
     * @param borderColor border color
     */
    public void applyBorder(BorderStyle borderStyle, XSSFColor borderColor) {
        this.cellStyle.setBorderTop(borderStyle);
        this.cellStyle.setBorderRight(borderStyle);
        this.cellStyle.setBorderBottom(borderStyle);
        this.cellStyle.setBorderLeft(borderStyle);

        this.cellStyle.setTopBorderColor(borderColor);
        this.cellStyle.setRightBorderColor(borderColor);
        this.cellStyle.setBottomBorderColor(borderColor);
        this.cellStyle.setLeftBorderColor(borderColor);
    }

    /**
     * Apply filters to header of spreadsheet.
     * @param sheet the spreadsheet used
     * @param lineIndex line number where filter will be applied
     * @param cellStart cell where the addition of the filter will begin
     * @param cellEnd cell where the addition of the filter will end
     */
    private void applyFilter(XSSFSheet sheet, int lineIndex, int cellStart, int cellEnd) {
        sheet.setAutoFilter(new CellRangeAddress(lineIndex, lineIndex, cellStart, cellEnd));
    }

    public void applyColorIn(int red, int green, int blue, ColorStyleFor colorStyleFor) {
        if (colorStyleFor.equals(ColorStyleFor.COLORCELL)) {
            setColorCell(red, green, blue);
        } else if (colorStyleFor.equals(ColorStyleFor.COLORBORDER)) {
            setColorBorder(red, green, blue);
        } else {
            LOGGER.error("No Informed ColorStyleFor");
        }
    }

    private void setColorCell(int red, int green, int blue) {
        this.colorCell = new XSSFColor(new Color(red, green, blue),null);
    }

    private void setColorBorder(int red, int green, int blue) {
        this.colorBorder = new XSSFColor(new Color(red, green, blue),null);
    }

    public XSSFCellStyle getCellStyle() {
        return cellStyle;
    }

    public XSSFFont getFontStyle() {
        return fontStyle;
    }

    public void setFontStyle(XSSFFont fontStyle) {
        this.fontStyle = fontStyle;
    }

    public XSSFColor getColorBorder() {
        return colorBorder;
    }

    public void setColorBorder(XSSFColor colorBorder) {
        this.colorBorder = colorBorder;
    }

    public XSSFColor getColorCell() {
        return colorCell;
    }

    public void setColorCell(XSSFColor colorCell) {
        this.colorCell = colorCell;
    }
}
