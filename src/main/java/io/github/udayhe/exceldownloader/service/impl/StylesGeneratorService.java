package io.github.udayhe.exceldownloader.service.impl;

import com.google.common.collect.ImmutableMap;
import io.github.udayhe.exceldownloader.enums.CustomStyle;
import io.github.udayhe.exceldownloader.service.IStyleGeneratorService;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Service;

import java.util.Map;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
@RequiredArgsConstructor
@Service
public class StylesGeneratorService implements IStyleGeneratorService {

    @Override
    public Map<CustomStyle, CellStyle> prepareStyles(Workbook workbook) {
        Font boldArial = createBoldArialFont(workbook);
        Font redBoldArial = createRedBoldArialFont(workbook);

        CellStyle rightAlignedStyle = createRightAlignedStyle(workbook);
        CellStyle greyCenteredBoldArialWithBorderStyle =
                createGreyCenteredBoldArialWithBorderStyle(workbook, boldArial);
        CellStyle redBoldArialWithBorderStyle =
                createRedBoldArialWithBorderStyle(workbook, redBoldArial);
        CellStyle rightAlignedDateFormatStyle =
                createRightAlignedDateFormatStyle(workbook);

        return ImmutableMap.of(CustomStyle.RIGHT_ALIGNED, rightAlignedStyle,
                CustomStyle.GREY_CENTERED_BOLD_ARIAL_WITH_BORDER, greyCenteredBoldArialWithBorderStyle,
                CustomStyle.RED_BOLD_ARIAL_WITH_BORDER, redBoldArialWithBorderStyle,
                CustomStyle.RIGHT_ALIGNED_DATE_FORMAT, rightAlignedDateFormatStyle);
    }

    private Font createBoldArialFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        return font;
    }

    private Font createRedBoldArialFont(Workbook workbook) {
        Font font = workbook.createFont();
        font.setFontName("Arial");
        font.setBold(true);
        font.setColor(IndexedColors.RED.index);
        return font;
    }

    private CellStyle createRightAlignedStyle(Workbook workBook) {
        CellStyle style = workBook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        return style;
    }

    private CellStyle createGreyCenteredBoldArialWithBorderStyle(Workbook workbook, Font boldArial) {
        CellStyle style = createBorderedStyle(workbook);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setFont(boldArial);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private CellStyle createRedBoldArialWithBorderStyle(Workbook workbook, Font redBoldArial) {
        CellStyle style = createBorderedStyle(workbook);
        style.setFont(redBoldArial);
        return style;
    }

    private CellStyle createRightAlignedDateFormatStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.RIGHT);
        style.setDataFormat((short) 14);
        return style;
    }

    private CellStyle createBorderedStyle(Workbook workbook) {
        BorderStyle thin = BorderStyle.THIN;
        short black = IndexedColors.BLACK.getIndex();
        CellStyle style = workbook.createCellStyle();
        style.setBorderRight(thin);
        style.setRightBorderColor(black);
        style.setBorderBottom(thin);
        style.setBottomBorderColor(black);
        style.setBorderLeft(thin);
        style.setLeftBorderColor(black);
        style.setBorderTop(thin);
        style.setTopBorderColor(black);
        return style;
    }
}
