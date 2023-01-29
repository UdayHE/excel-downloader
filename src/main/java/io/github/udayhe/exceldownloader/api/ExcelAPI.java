package io.github.udayhe.exceldownloader.api;

import io.github.udayhe.exceldownloader.service.IExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

/**
 * @author udayhegde
 * @Date 22/01/23
 */
@RestController
@RequiredArgsConstructor
@RequestMapping("/excel")
public class ExcelAPI {

    private final IExcelService excelService;


    @PostMapping("/xlsx")
    public ResponseEntity<String> generateXlsx() throws IOException {
        return ResponseEntity.ok(excelService.generateXlsxReport());
    }

    @PostMapping("/xls")
    public ResponseEntity<String> generateXls() throws IOException {
        return ResponseEntity.ok(excelService.generateXlsReport());
    }


    private ResponseEntity<byte[]> createResponseEntity(byte[] report, String fileName) {
        return ResponseEntity.ok()
                .contentType(MediaType.APPLICATION_OCTET_STREAM)
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=" + fileName)
                .body(report);
    }
}
