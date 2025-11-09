package com.javaweb.controller;

import com.javaweb.dto.CouncilEvaluationDto;
import com.javaweb.service.CouncilEvaluationExportService;
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
    private final CouncilEvaluationExportService councilSvc;

    public ExportController(CouncilEvaluationExportService councilSvc) {
        this.councilSvc = councilSvc;
    }

    @PostMapping("/xlsx")
    public ResponseEntity<byte[]> council(@RequestBody CouncilEvaluationDto.Root payload) throws Exception {
        Workbook wb = councilSvc.buildWorkbook(payload);
        return buildResponse(wb, "phieu_cham_hoi_dong.xlsx");
    }

    private ResponseEntity<byte[]> buildResponse(Workbook workbook, String filename) throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        workbook.write(bos);
        workbook.close();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDisposition(ContentDisposition.attachment().filename(filename).build());
        return new ResponseEntity<>(bos.toByteArray(), headers, HttpStatus.OK);
    }
}

