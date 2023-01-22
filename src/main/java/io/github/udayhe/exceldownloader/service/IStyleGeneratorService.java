package io.github.udayhe.exceldownloader.service;

import io.github.udayhe.exceldownloader.enums.CustomStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Map;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
public interface IStyleGeneratorService {

    Map<CustomStyle, CellStyle> prepareStyles(Workbook workbook);
}
