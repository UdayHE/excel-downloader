package io.github.udayhe.exceldownloader.service.impl;

import io.github.udayhe.exceldownloader.enums.CustomStyle;
import io.github.udayhe.exceldownloader.service.IExcelService;
import io.github.udayhe.exceldownloader.service.IStyleGeneratorService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Map;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
@RequiredArgsConstructor
@Service
public class ExcelService implements IExcelService {

    private final IStyleGeneratorService stylesGeneratorService;

    @Override
    public byte[] generateXlsxReport() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        return generateReport(workbook);
    }

    @Override
    public byte[] generateXlsReport() throws IOException {
        Workbook wb = new HSSFWorkbook();

        return generateReport(wb);
    }


    private byte[] generateReport(Workbook workbook) throws IOException {
        Map<CustomStyle, CellStyle> styles = stylesGeneratorService.prepareStyles(workbook);
        Sheet sheet = workbook.createSheet("Example sheet name");

        setColumnsWidth(sheet);

        createHeaderRow(sheet, styles);
        createStringsRow(sheet, styles);
        createDoublesRow(sheet, styles);
        createDatesRow(sheet, styles);

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);

        out.close();
        workbook.close();

        return out.toByteArray();
    }

    private void setColumnsWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 256 * 20);
        for (int i = 0; i < 5; i++) {
            sheet.setColumnWidth(i, 256 * 15);
        }
    }

    private void createHeaderRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(0);

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("Column $i");
            cell.setCellStyle(styles.get(CustomStyle.GREY_CENTERED_BOLD_ARIAL_WITH_BORDER));
        }
    }

    private void createRowLabelCell(Row row, Map<CustomStyle, CellStyle> styles, String label) {
        Cell rowLabel = row.createCell(0);
        rowLabel.setCellValue(label);
        rowLabel.setCellStyle(styles.get(CustomStyle.RED_BOLD_ARIAL_WITH_BORDER));
    }

    private void createStringsRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(1);
        createRowLabelCell(row, styles, "Strings row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("String $i");
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED));
        }

    }

    private void createDoublesRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(2);
        createRowLabelCell(row, styles, "Doubles row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(new BigDecimal(i + ".99").doubleValue());
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED));
        }
    }

    private void createDatesRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(3);
        createRowLabelCell(row, styles, "Dates row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue((LocalDate.now()));
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED_DATE_FORMAT));
        }
    }
}