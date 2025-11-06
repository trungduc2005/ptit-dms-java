package com.javaweb.controller;

import com.javaweb.dto.EvaluationDto;
import com.javaweb.service.EvaluationExportService;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.http.*;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayOutputStream;

@RestController
@RequestMapping("/api/export")
public class ExportController {
    private final EvaluationExportService svc;
    public ExportController(EvaluationExportService svc){ this.svc = svc; }
    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> xlsx(@RequestBody EvaluationDto.Root payload) throws Exception {
        Workbook wb = svc.buildWorkbook(payload);
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        wb.write(bos); wb.close();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDisposition(ContentDisposition.attachment().filename("phieu_cham.xlsx").build());
        return new ResponseEntity<>(bos.toByteArray(), headers, HttpStatus.OK);
    }
}

