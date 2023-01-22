package io.github.udayhe.exceldownloader.service;

import java.io.IOException;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
public interface IExcelService {

    byte[] generateXlsxReport() throws IOException;

    byte[] generateXlsReport() throws IOException;
}
