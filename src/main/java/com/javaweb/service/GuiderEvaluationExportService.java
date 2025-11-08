package com.javaweb.service;

import com.javaweb.dto.GuiderEvaluationDto;
import com.javaweb.dto.GuiderEvaluationDto.EvaluationForm;
import com.javaweb.dto.GuiderEvaluationDto.Indicator;
import com.javaweb.dto.GuiderEvaluationDto.Pi;
import com.javaweb.dto.GuiderEvaluationDto.Score;
import com.javaweb.dto.GuiderEvaluationDto.Student;
import com.javaweb.dto.GuiderEvaluationDto.StudentEvaluation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.Collections;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class GuiderEvaluationExportService {

    private static final int BASE_COLUMNS = 6;

    public Workbook buildWorkbook(GuiderEvaluationDto.Root root) {
        Workbook workbook = new XSSFWorkbook();
        Styles styles = new Styles(workbook);

        List<EvaluationForm> forms = root != null && root.getEvaluationForm() != null
                ? root.getEvaluationForm()
                : Collections.emptyList();
        List<Student> students = root != null && root.getStudents() != null
                ? root.getStudents()
                : Collections.emptyList();

        List<FormBlock> blocks = buildBlocks(forms, students, styles, workbook);

        Sheet sheet = workbook.createSheet("Huong dan");
        configureColumns(sheet, blocks);

        int rowIndex = 0;
        rowIndex = buildHeader(sheet, rowIndex, blocks, styles);
        populateRows(sheet, rowIndex, students, blocks, styles);

        return workbook;
    }

    private List<FormBlock> buildBlocks(List<EvaluationForm> forms,
                                        List<Student> students,
                                        Styles styles,
                                        Workbook workbook) {
        List<FormBlock> blocks = new ArrayList<>();
        Map<String, EvaluationTemplate> templates = buildEvaluationTemplates(students);
        short[] palette = {
                IndexedColors.ROSE.getIndex(),
                IndexedColors.LIGHT_GREEN.getIndex(),
                IndexedColors.LAVENDER.getIndex(),
                IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex()
        };

        List<String> covered = new ArrayList<>();

        if (forms != null) {
            for (int i = 0; i < forms.size(); i++) {
                EvaluationForm form = forms.get(i);
                List<PiEntry> piEntries = expandPis(form);
                if (piEntries.isEmpty()) {
                    EvaluationTemplate template = form != null ? templates.get(form.getEvaluationId()) : null;
                    if (template != null) {
                        piEntries = new ArrayList<>(template.piEntries());
                    }
                }
                if (piEntries.isEmpty()) {
                    piEntries.add(new PiEntry("N/A", "Chưa cấu hình PI"));
                }
                short color = palette[i % palette.length];
                CellStyle blockHeader = createColoredStyle(workbook, styles.header, color);
                CellStyle blockCell = createColoredStyle(workbook, styles.cellCenter, color);
                blocks.add(new FormBlock(form, piEntries, blockHeader, blockCell));
                if (form != null && form.getEvaluationId() != null) {
                    covered.add(form.getEvaluationId());
                }
            }
        }

        int offset = blocks.size();
        for (EvaluationTemplate template : templates.values()) {
            if (template.evaluationId() != null && covered.contains(template.evaluationId())) {
                continue;
            }
            List<PiEntry> entries = template.piEntries().isEmpty()
                    ? Collections.singletonList(new PiEntry("N/A", "Chưa cấu hình PI"))
                    : template.piEntries();
            EvaluationForm fallback = new EvaluationForm();
            fallback.setEvaluationId(template.evaluationId());
            fallback.setTitle(template.title());

            short color = palette[(offset++) % palette.length];
            CellStyle blockHeader = createColoredStyle(workbook, styles.header, color);
            CellStyle blockCell = createColoredStyle(workbook, styles.cellCenter, color);
            blocks.add(new FormBlock(fallback, new ArrayList<>(entries), blockHeader, blockCell));
        }
        return blocks;
    }

    private List<PiEntry> expandPis(EvaluationForm form) {
        List<PiEntry> entries = new ArrayList<>();
        if (form == null || form.getIndicators() == null) {
            return entries;
        }
        for (Indicator indicator : form.getIndicators()) {
            List<Pi> pis = indicator != null && indicator.getPis() != null
                    ? indicator.getPis()
                    : Collections.emptyList();
            if (pis.isEmpty()) {
                String fallbackId = indicator != null && indicator.getCloId() != null
                        ? indicator.getCloId()
                        : "PI";
                String fallbackDesc = indicator != null && indicator.getCloDescription() != null
                        ? indicator.getCloDescription()
                        : "";
                entries.add(new PiEntry(fallbackId, fallbackDesc));
            } else {
                for (Pi pi : pis) {
                    String label = (pi.getCloPisId() != null ? pi.getCloPisId() : "PI")
                            + (pi.getCloPisDescription() != null ? ": " + pi.getCloPisDescription() : "");
                    entries.add(new PiEntry(pi.getCloPisId(), label));
                }
            }
        }
        return entries;
    }

    private void configureColumns(Sheet sheet, List<FormBlock> blocks) {
        sheet.setColumnWidth(0, 5 * 256);
        sheet.setColumnWidth(1, 16 * 256);
        sheet.setColumnWidth(2, 24 * 256);
        sheet.setColumnWidth(3, 16 * 256);
        sheet.setColumnWidth(4, 26 * 256);
        sheet.setColumnWidth(5, 42 * 256);

        int column = BASE_COLUMNS;
        for (FormBlock block : blocks) {
            for (int i = 0; i < block.piEntries().size(); i++) {
                sheet.setColumnWidth(column++, 22 * 256);
            }
        }
    }

    private int buildHeader(Sheet sheet,
                            int startRow,
                            List<FormBlock> blocks,
                            Styles styles) {

        Row row0 = sheet.createRow(startRow);
        Row row1 = sheet.createRow(startRow + 1);
        Row row2 = sheet.createRow(startRow + 2);

                String[] baseHeaders = {
                "STT",
                "Mã sinh viên",
                "Họ và tên đệm",
                "Tên",
                "Giáo viên hướng dẫn",
                "Tên đề tài đồ án tốt nghiệp"
        };

        for (int i = 0; i < baseHeaders.length; i++) {
            merge(sheet, startRow, startRow + 2, i, i);
            setCell(row0, i, baseHeaders[i], styles.header);
            setCell(row1, i, "", styles.header);
            setCell(row2, i, "", styles.header);
        }

        int columnIndex = baseHeaders.length;
        for (FormBlock block : blocks) {
            int blockSize = block.piEntries().size();
            int blockStart = columnIndex;
            int blockEnd = columnIndex + blockSize - 1;

            merge(sheet, row0.getRowNum(), row0.getRowNum(), blockStart, blockEnd);
            setCell(row0, blockStart, blockTitle(block.form()), block.headerStyle());

            merge(sheet, row1.getRowNum(), row1.getRowNum(), blockStart, blockEnd);
            setCell(row1, blockStart, blockSubtitle(block.form()), styles.subHeader);

            for (PiEntry entry : block.piEntries()) {
                setCell(row2, columnIndex++, entry.label(), block.headerStyle());
            }
        }

        return startRow + 3;
    }

    private void populateRows(Sheet sheet,
                              int startRow,
                              List<Student> students,
                              List<FormBlock> blocks,
                              Styles styles) {
        int rowIndex = startRow;
        int order = 1;

        for (Student student : students) {
            Row row = sheet.createRow(rowIndex++);
            setCell(row, 0, order++, styles.cellCenter);
            setCell(row, 1, nullSafe(student.getStudentId()), styles.cellCenter);
            String[] nameParts = splitName(student.getStudentName());
            setCell(row, 2, nameParts[0], styles.cellLeft);
            setCell(row, 3, nameParts[1], styles.cellLeft);
            setCell(row, 4, nullSafe(student.getGuiderName()), styles.cellLeftWrap);
            setCell(row, 5, nullSafe(student.getProjectName()), styles.cellLeftWrap);

            Map<String, StudentEvaluation> evaluationMap = student != null
                    ? student.evaluationMap()
                    : Collections.emptyMap();

            int columnIndex = BASE_COLUMNS;
            for (FormBlock block : blocks) {
                StudentEvaluation evaluation = evaluationMap.get(
                        block.form() != null ? block.form().getEvaluationId() : null);
                Map<String, Double> scoreMap = evaluation != null ? evaluation.scoreMap() : Collections.emptyMap();
                for (PiEntry entry : block.piEntries()) {
                    Double value = scoreMap.get(entry.piId());
                    setCell(row, columnIndex++, value, block.cellStyle());
                }
            }
        }
    }

    private String blockTitle(EvaluationForm form) {
        if (form == null) {
            return "Phiếu đánh giá";
        }
        String weekPart = form.getReportWeek() != null ? "Đợt " + form.getReportWeek() : null;
        if (weekPart != null && form.getTitle() != null) {
            return form.getTitle();
        }
        if (form.getTitle() != null) {
            return form.getTitle();
        }
        return weekPart != null ? weekPart : "Phiếu đánh giá";
    }

    private String blockSubtitle(EvaluationForm form) {
        if (form == null) {
            return "";
        }
        java.util.List<String> parts = new java.util.ArrayList<>();
        if (form.getReportWeek() != null) {
            parts.add("Tuần " + form.getReportWeek());
        }
        if (form.getAcademicYear() != null) {
            parts.add("Năm học " + form.getAcademicYear());
        }
        if (form.getDescription() != null) {
            parts.add(form.getDescription());
        }
        return String.join(" - ", parts);
    }
private CellStyle createColoredStyle(Workbook workbook, CellStyle template, short color) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(template);
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }

    private void merge(Sheet sheet, int firstRow, int lastRow, int firstCol, int lastCol) {
        if (firstRow > lastRow || firstCol > lastCol) {
            return;
        }
        if (firstRow == lastRow && firstCol == lastCol) {
            return;
        }
        sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow, firstCol, lastCol));
    }

    private void setCell(Row row, int columnIndex, Object value, CellStyle style) {
        if (row == null) {
            return;
        }
        Cell cell = row.getCell(columnIndex);
        if (cell == null) {
            cell = row.createCell(columnIndex);
        }
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof Number) {
            cell.setCellValue(((Number) value).doubleValue());
        } else {
            cell.setCellValue(String.valueOf(value));
        }
        if (style != null) {
            cell.setCellStyle(style);
        }
    }

    private String[] splitName(String fullName) {
        if (fullName == null || fullName.isBlank()) {
            return new String[]{"", ""};
        }
        String normalized = fullName.trim().replaceAll("\\s+", " ");
        int lastSpace = normalized.lastIndexOf(' ');
        if (lastSpace < 0) {
            return new String[]{"", normalized};
        }
        return new String[]{
                normalized.substring(0, lastSpace),
                normalized.substring(lastSpace + 1)
        };
    }

    private String nullSafe(String value) {
        return value != null ? value : "";
    }

    private record PiEntry(String piId, String label) {}

    private record FormBlock(EvaluationForm form,
                             List<PiEntry> piEntries,
                             CellStyle headerStyle,
                             CellStyle cellStyle) {}

    private record EvaluationTemplate(String evaluationId,
                                      String title,
                                      List<PiEntry> piEntries) {}

    private Map<String, EvaluationTemplate> buildEvaluationTemplates(List<Student> students) {
        Map<String, EvaluationTemplate> templates = new LinkedHashMap<>();
        if (students == null) {
            return templates;
        }
        for (Student student : students) {
            if (student.getEvaluations() == null) {
                continue;
            }
            for (StudentEvaluation evaluation : student.getEvaluations()) {
                String evalId = evaluation.getEvaluationId();
                if (evalId == null || templates.containsKey(evalId)) {
                    continue;
                }
                List<PiEntry> entries = new ArrayList<>();
                if (evaluation.getScores() != null) {
                    for (Score score : evaluation.getScores()) {
                        String label = score.getPiId() != null ? score.getPiId() : "PI";
                        entries.add(new PiEntry(score.getPiId(), label));
                    }
                }
                templates.put(evalId, new EvaluationTemplate(
                        evalId,
                        evaluation.getEvaluationTitle(),
                        entries));
            }
        }
        return templates;
    }


    private static class Styles {
        final CellStyle header;
        final CellStyle subHeader;
        final CellStyle cellCenter;
        final CellStyle cellLeft;
        final CellStyle cellLeftWrap;

        Styles(Workbook workbook) {
            Font normal = workbook.createFont();
            normal.setFontName("Times New Roman");
            normal.setFontHeightInPoints((short) 12);

            Font bold = workbook.createFont();
            bold.setBold(true);
            bold.setFontName("Times New Roman");
            bold.setFontHeightInPoints((short) 12);

            Font italic = workbook.createFont();
            italic.setItalic(true);
            italic.setFontName("Times New Roman");
            italic.setFontHeightInPoints((short) 12);

            header = workbook.createCellStyle();
            header.setFont(bold);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);
            addBorder(header);

            subHeader = workbook.createCellStyle();
            subHeader.setFont(italic);
            subHeader.setAlignment(HorizontalAlignment.CENTER);
            subHeader.setVerticalAlignment(VerticalAlignment.CENTER);
            subHeader.setWrapText(true);
            addBorder(subHeader);

            cellCenter = workbook.createCellStyle();
            cellCenter.setFont(normal);
            cellCenter.setAlignment(HorizontalAlignment.CENTER);
            cellCenter.setVerticalAlignment(VerticalAlignment.CENTER);
            addBorder(cellCenter);

            cellLeft = workbook.createCellStyle();
            cellLeft.setFont(normal);
            cellLeft.setAlignment(HorizontalAlignment.LEFT);
            cellLeft.setVerticalAlignment(VerticalAlignment.CENTER);
            addBorder(cellLeft);

            cellLeftWrap = workbook.createCellStyle();
            cellLeftWrap.cloneStyleFrom(cellLeft);
            cellLeftWrap.setWrapText(true);
        }

        private void addBorder(CellStyle style) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
        }
    }
}
