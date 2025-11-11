package com.javaweb.service;

import com.javaweb.dto.GuiderEvaluationDto;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.springframework.stereotype.Service;

import java.util.Collections;

@Service
public class ReviewerEvaluationExportService extends GuiderEvaluationExportService {

    private static final SheetLayout REVIEWER_SHEET =
            new SheetLayout("PB_CaNhan", PB_COLUMNS, styles ->
                    Collections.singletonList(new ExtraColumn(
                            "Nh\u1eadn x\u00e9t/Y\u00eau c\u1ea7u s\u1eeda \u0111\u1ed5i (\u1ebfu c\u00f3)",
                            32 * 256,
                            styles.header,
                            styles.cellLeftWrap,
                            student -> ""
                    )));

    public Workbook buildWorkbook(GuiderEvaluationDto.Root root) {
        return buildWorkbook(root, REVIEWER_SHEET);
    }

    @Override
    protected int headerStartRow() {
        return 2;
    }

    @Override
    protected short[] blockPalette() {
        return new short[]{IndexedColors.LAVENDER.getIndex()};
    }
}
