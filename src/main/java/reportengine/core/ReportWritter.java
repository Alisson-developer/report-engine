package reportengine.core;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public class ReportWritter<T> {
    private static final Logger LOGGER = LoggerFactory.getLogger(ReportWritter.class);
    private final XSSFWorkbook workbook = new XSSFWorkbook();
    private final XSSFSheet sheet = workbook.createSheet("Relatorio-Batatinhas");

    public ReportWritter() {
    }

    public void writeInXlsx(T obj, List<T> data) throws IOException {
        List<List<Object>> registers = new ArrayList<>();
        Field[] fields = obj.getClass().getDeclaredFields();
        buildContent(data, registers, fields);
        putData(registers, fields);
    }

    protected void buildContent(List<T> objs, List<List<Object>> registers, Field[] fields) {
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
            attributeFields.add(fieldName);
        }
    }

    private void makeDataValues(List<T> objs, List<Object> attributeValues, Field[] fields) {
        objs.forEach(obj -> {
            for(Field field: fields) {
                field.setAccessible(true);
                try {
                    attributeValues.add(field.get(obj));
                } catch (IllegalAccessException e) {
                    LOGGER.error(e.getMessage());
                }
            }
        });
    }

    private void putData(List<List<Object>> registers, Field[] fields) throws IOException {
        int rowCount = 0;
        int rowTitle = 0;

        for(List<Object> register: registers) {
            XSSFRow row = sheet.createRow(rowCount++);
            int cellCount = 0;

            for(Object value: register) {
                if(cellCount >= fields.length) {
                    cellCount = 0;
                    row = sheet.createRow(rowCount++);
                }

                XSSFCell cell = row.createCell(cellCount++);
                cell.setCellValue(value.toString());

                if(rowCount == 1) {
                    applyTitleStyle(sheet, rowTitle, cellCount);
                }
            }

        }

        File pathXlsx = new File("src/main/resources/static/relatorio.xlsx");
        FileOutputStream outPut = new FileOutputStream(pathXlsx);
        workbook.write(outPut);
        outPut.close();
    }

    private void applyTitleStyle(XSSFSheet sheet, int lineIndex, int cellIndex) {
        XSSFColor colorCell = new XSSFColor(new Color(43,150,150), null);
        XSSFColor colorBorder = new XSSFColor(new Color(50,50,50), null);
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        XSSFFont font = workbook.createFont();
        font.setBold(true);

        cellStyle.setFillForegroundColor(colorCell);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setFont(font);

        buildBorder(cellStyle, BorderStyle.THIN, colorBorder);

        XSSFCell cell = sheet.getRow(lineIndex).getCell(--cellIndex);
        cell.setCellStyle(cellStyle);
    }

    private void buildBorder(XSSFCellStyle cellStyle, BorderStyle borderStyle, XSSFColor borderColor) {
        cellStyle.setBorderTop(borderStyle);
        cellStyle.setBorderRight(borderStyle);
        cellStyle.setBorderBottom(borderStyle);
        cellStyle.setBorderLeft(borderStyle);

        cellStyle.setTopBorderColor(borderColor);
        cellStyle.setRightBorderColor(borderColor);
        cellStyle.setBottomBorderColor(borderColor);
        cellStyle.setLeftBorderColor(borderColor);
    }
}