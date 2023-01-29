package io.github.udayhe.exceldownloader.service.impl;

import io.github.udayhe.exceldownloader.enums.CustomStyle;
import io.github.udayhe.exceldownloader.helper.ExcelHelper;
import io.github.udayhe.exceldownloader.service.IExcelService;
import io.github.udayhe.exceldownloader.service.IStyleGeneratorService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.util.Map;

import static io.github.udayhe.exceldownloader.constant.Constant.XLSX_PATH;
import static io.github.udayhe.exceldownloader.constant.Constant.XLS_PATH;
import static org.apache.commons.io.FileUtils.writeByteArrayToFile;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
@RequiredArgsConstructor
@Service
public class ExcelService implements IExcelService {

    private final IStyleGeneratorService stylesGeneratorService;
    private final ExcelHelper excelHelper;

    @Override
    public String generateXlsxReport() throws IOException {
        Workbook workbook = new XSSFWorkbook();
        writeByteArrayToFile(new File(XLSX_PATH), generateReport(workbook));
        return XLSX_PATH;
    }

    @Override
    public String generateXlsReport() throws IOException {
        Workbook workbook = new HSSFWorkbook();
        writeByteArrayToFile(new File(XLS_PATH), generateReport(workbook));
        return XLS_PATH;
    }

    public byte[] generateReport(Workbook workbook) throws IOException {
        Map<CustomStyle, CellStyle> styles = stylesGeneratorService.prepareStyles(workbook);
        Sheet sheet = workbook.createSheet("Example sheet name");
        prepareExcelSheet(sheet, styles);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        out.close();
        workbook.close();
        return out.toByteArray();
    }

    private void prepareExcelSheet(Sheet sheet, Map<CustomStyle, CellStyle> styles) {
        excelHelper.setColumnsWidth(sheet);
        excelHelper.createHeaderRow(sheet, styles);
        excelHelper.createStringsRow(sheet, styles);
        excelHelper.createDoublesRow(sheet, styles);
        excelHelper.createDatesRow(sheet, styles);
    }
}