package io.github.udayhe.exceldownloader.service;

import java.io.IOException;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
public interface IExcelService {

    String generateXlsxReport() throws IOException;

    String generateXlsReport() throws IOException;
}
