package io.github.udayhe.exceldownloader.helper;

import io.github.udayhe.exceldownloader.enums.CustomStyle;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Component;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Map;

/**
 * @author udayhegde
 * @Date 29/01/23
 */
@Component
@RequiredArgsConstructor
public class ExcelHelper {


    public void setColumnsWidth(Sheet sheet) {
        sheet.setColumnWidth(0, 256 * 20);
        for (int i = 0; i < 5; i++) {
            sheet.setColumnWidth(i, 256 * 15);
        }
    }

    public void createHeaderRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(0);

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("Column " + i);
            cell.setCellStyle(styles.get(CustomStyle.GREY_CENTERED_BOLD_ARIAL_WITH_BORDER));
        }
    }

    public void createRowLabelCell(Row row, Map<CustomStyle, CellStyle> styles, String label) {
        Cell rowLabel = row.createCell(0);
        rowLabel.setCellValue(label);
        rowLabel.setCellStyle(styles.get(CustomStyle.RED_BOLD_ARIAL_WITH_BORDER));
    }

    public void createStringsRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(1);
        createRowLabelCell(row, styles, "Strings row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue("String " + i);
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED));
        }

    }

    public void createDoublesRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(2);
        createRowLabelCell(row, styles, "Doubles row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(new BigDecimal(i + ".99").doubleValue());
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED));
        }
    }

    public void createDatesRow(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        Row row = sheet.createRow(3);
        createRowLabelCell(row, styles, "Dates row");

        for (int i = 0; i < 5; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue((LocalDate.now()));
            cell.setCellStyle(styles.get(CustomStyle.RIGHT_ALIGNED_DATE_FORMAT));
        }
    }
}
