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

    private static final ColumnType[] GVHD_COLUMNS = {
            ColumnType.ORDER,
            ColumnType.STUDENT_ID,
            ColumnType.STUDENT_LAST,
            ColumnType.STUDENT_FIRST,
            ColumnType.GUIDER,
            ColumnType.PROJECT
    };

    protected static final ColumnType[] PB_COLUMNS = {
            ColumnType.ORDER,
            ColumnType.STUDENT_ID,
            ColumnType.STUDENT_LAST,
            ColumnType.STUDENT_FIRST,
            ColumnType.GUIDER,
            ColumnType.PROJECT,
            ColumnType.REVIEWER
    };

    private static final SheetLayout GVHD_SHEET =
            new SheetLayout("GVHD_CaNhan", GVHD_COLUMNS, styles -> Collections.emptyList());

    public Workbook buildWorkbook(GuiderEvaluationDto.Root root) {
        return buildWorkbook(root, GVHD_SHEET);
    }

    protected Workbook buildWorkbook(GuiderEvaluationDto.Root root, SheetLayout sheetLayout) {
        Workbook workbook = new XSSFWorkbook();
        Styles styles = new Styles(workbook);

        List<EvaluationForm> forms = root != null && root.getEvaluationForm() != null
                ? root.getEvaluationForm()
                : Collections.emptyList();
        List<Student> students = root != null && root.getStudents() != null
                ? root.getStudents()
                : Collections.emptyList();

        List<FormBlock> blocks = buildBlocks(forms, students, styles, workbook);
        List<ExtraColumn> extraColumns = sheetLayout != null && sheetLayout.extraColumnBuilder() != null
                ? sheetLayout.extraColumnBuilder().build(styles)
                : Collections.emptyList();

        String sheetName = sheetLayout != null ? sheetLayout.sheetName() : "Export";
        ColumnType[] layout = sheetLayout != null ? sheetLayout.columns() : GVHD_COLUMNS;

        buildSheet(workbook, sheetName, layout, students, blocks, extraColumns, styles);

        if (workbook.getNumberOfSheets() == 0) {
            workbook.createSheet("Empty");
        }

        return workbook;
    }

    /**
     * Starting row index for the header (allows subclasses to prepend blank rows).
     */
    protected int headerStartRow() {
        return 0;
    }

    private void buildSheet(Workbook workbook,
                            String sheetName,
                            ColumnType[] layout,
                            List<Student> students,
                            List<FormBlock> blocks,
                            List<ExtraColumn> extras,
                            Styles styles) {

        Sheet sheet = workbook.createSheet(sheetName);
        configureColumns(sheet, layout, blocks, extras);

        int headerRowIndex = headerStartRow();
        if (headerRowIndex > 0) {
            for (int i = 0; i < headerRowIndex; i++) {
                if (sheet.getRow(i) == null) {
                    sheet.createRow(i);
                }
            }
        }

        int rowIndex = buildHeader(sheet, headerRowIndex, layout, blocks, extras, styles);
        populateRows(sheet, rowIndex, students, layout, blocks, extras, styles);
    }

    private void configureColumns(Sheet sheet,
                                  ColumnType[] layout,
                                  List<FormBlock> blocks,
                                  List<ExtraColumn> extras) {
        for (int i = 0; i < layout.length; i++) {
            sheet.setColumnWidth(i, layout[i].width());
        }

        int column = layout.length;
        for (FormBlock block : blocks) {
            for (int i = 0; i < block.piEntries().size(); i++) {
                sheet.setColumnWidth(column++, 22 * 256);
            }
        }

        if (extras != null) {
            for (ExtraColumn extra : extras) {
                sheet.setColumnWidth(column++, extra.width());
            }
        }
    }

    private int buildHeader(Sheet sheet,
                            int startRow,
                            ColumnType[] layout,
                            List<FormBlock> blocks,
                            List<ExtraColumn> extras,
                            Styles styles) {

        Row row0 = sheet.createRow(startRow);
        Row row2 = sheet.createRow(startRow + 1);
        Row row3 = sheet.createRow(startRow + 2);

        for (int i = 0; i < layout.length; i++) {
            merge(sheet, startRow, startRow + 2, i, i);
            setCell(row0, i, layout[i].header(), styles.header);
            setCell(row2, i, "", styles.header);
            setCell(row3, i, "", styles.header);
        }

        int columnIndex = layout.length;
        for (FormBlock block : blocks) {
            int blockSize = block.piEntries().size();
            int blockStart = columnIndex;
            int blockEnd = columnIndex + blockSize - 1;

            merge(sheet, row0.getRowNum(), row0.getRowNum(), blockStart, blockEnd);
            applyHorizontalBorder(sheet, row0.getRowNum(), blockStart, blockEnd, block.headerStyle());
            setCell(row0, blockStart, blockTitle(block.form()), block.headerStyle());

            String currentIndicator = null;
            int indicatorStart = columnIndex;
            for (PiEntry entry : block.piEntries()) {
                if (currentIndicator == null) {
                    currentIndicator = entry.indicatorLabel();
                } else if (!currentIndicator.equals(entry.indicatorLabel())) {
                    setIndicatorHeader(sheet, row2, indicatorStart, columnIndex - 1, currentIndicator, block.headerStyle());
                    indicatorStart = columnIndex;
                    currentIndicator = entry.indicatorLabel();
                }
                setCell(row3, columnIndex++, entry.label(), block.headerStyle());
            }
            if (currentIndicator != null) {
                setIndicatorHeader(sheet, row2, indicatorStart, columnIndex - 1, currentIndicator, block.headerStyle());
            }
        }

        if (extras != null) {
            for (ExtraColumn extra : extras) {
                CellStyle headerStyle = extra.headerStyle() != null ? extra.headerStyle() : styles.header;
                merge(sheet, startRow, startRow + 2, columnIndex, columnIndex);
                setCell(row0, columnIndex, extra.header(), headerStyle);
                setCell(row2, columnIndex, "", headerStyle);
                setCell(row3, columnIndex, "", headerStyle);
                columnIndex++;
            }
        }

        return startRow + 3;
    }

    private void populateRows(Sheet sheet,
                              int startRow,
                              List<Student> students,
                              ColumnType[] layout,
                              List<FormBlock> blocks,
                              List<ExtraColumn> extras,
                              Styles styles) {

        int rowIndex = startRow;
        int order = 1;

        for (Student student : students) {
            Row row = sheet.createRow(rowIndex++);
            String[] nameParts = splitName(student.getStudentName());
            int baseColumnIndex = 0;
            for (ColumnType column : layout) {
                switch (column) {
                    case ORDER -> setCell(row, baseColumnIndex, order, styles.cellCenter);
                    case STUDENT_ID -> setCell(row, baseColumnIndex, nullSafe(student.getStudentId()), styles.cellCenter);
                    case STUDENT_LAST -> setCell(row, baseColumnIndex, nameParts[0], styles.cellLeft);
                    case STUDENT_FIRST -> setCell(row, baseColumnIndex, nameParts[1], styles.cellLeft);
                    case GUIDER -> setCell(row, baseColumnIndex, nullSafe(student.getGuiderName()), styles.cellLeftWrap);
                    case PROJECT -> setCell(row, baseColumnIndex, nullSafe(student.getProjectName()), styles.cellLeftWrap);
                    case REVIEWER -> {
                        setCell(row, baseColumnIndex, "", styles.cellLeftWrap);
                    }
                }

                if (column == ColumnType.ORDER) {
                    order++;
                }
                baseColumnIndex++;
            }

            Map<String, StudentEvaluation> evaluationMap = student != null
                    ? student.evaluationMap()
                    : Collections.emptyMap();

            int columnIndex = layout.length;
            for (FormBlock block : blocks) {
                StudentEvaluation evaluation = evaluationMap.get(
                        block.form() != null ? block.form().getEvaluationId() : null);
                Map<String, Double> scoreMap = evaluation != null ? evaluation.scoreMap() : Collections.emptyMap();

                for (PiEntry entry : block.piEntries()) {
                    Double value = scoreMap.get(entry.piId());
                    setCell(row, columnIndex++, value, block.cellStyle());
                }
            }

            if (extras != null) {
                for (ExtraColumn extra : extras) {
                    Object value = extra.valueProvider() != null ? extra.valueProvider().apply(student) : "";
                    setCell(row, columnIndex++, value, extra.cellStyle());
                }
            }
        }
    }

    private List<FormBlock> buildBlocks(List<EvaluationForm> forms,
                                        List<Student> students,
                                        Styles styles,
                                        Workbook workbook) {
        List<FormBlock> blocks = new ArrayList<>();
        Map<String, EvaluationTemplate> templates = buildEvaluationTemplates(students);
        short[] palette = blockPalette();
        if (palette == null || palette.length == 0) {
            palette = new short[]{IndexedColors.ROSE.getIndex()};
        }

        List<String> covered = new ArrayList<>();

        if (forms != null) {
            for (int i = 0; i < forms.size(); i++) {
                EvaluationForm form = forms.get(i);
                List<PiEntry> piEntries = new ArrayList<>(expandPis(form));
                if (piEntries.isEmpty()) {
                    EvaluationTemplate template = form != null ? templates.get(form.getEvaluationId()) : null;
                    if (template != null) {
                        piEntries = new ArrayList<>(template.piEntries());
                    }
                }
                if (piEntries.isEmpty()) {
                    piEntries.add(new PiEntry("N/A", "Ch\u00c3\u0192\u00c6\u2019\u00c3\u00a2\u00e2\u201a\u00ac\u00c2\u00a0\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00b0a c\u00c3\u0192\u00c6\u2019\u00c3\u201a\u00c2\u00a1\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00ba\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00a5u h\u00c3\u0192\u00c6\u2019\u00c3\u2020\u00e2\u20ac\u2122\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00acnh PI", "CLO"));
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
                    ? Collections.singletonList(new PiEntry("N/A", "Ch\u00c3\u0192\u00c6\u2019\u00c3\u00a2\u00e2\u201a\u00ac\u00c2\u00a0\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00b0a c\u00c3\u0192\u00c6\u2019\u00c3\u201a\u00c2\u00a1\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00ba\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00a5u h\u00c3\u0192\u00c6\u2019\u00c3\u2020\u00e2\u20ac\u2122\u00c3\u0192\u00e2\u20ac\u0161\u00c3\u201a\u00c2\u00acnh PI", "CLO"))
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

    protected short[] blockPalette() {
        return new short[]{
                IndexedColors.ROSE.getIndex(),
                IndexedColors.LIGHT_GREEN.getIndex(),
                IndexedColors.LAVENDER.getIndex(),
                IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex()
        };
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
            String indicatorLabel = buildIndicatorLabel(indicator);
            if (pis.isEmpty()) {
                String fallbackId = indicator != null && indicator.getCloId() != null
                        ? indicator.getCloId()
                        : "PI";
                String fallbackDesc = indicator != null && indicator.getCloDescription() != null
                        ? indicator.getCloDescription()
                        : "";
                entries.add(new PiEntry(fallbackId, fallbackDesc, indicatorLabel));
            } else {
                for (Pi pi : pis) {
                    String label = (pi.getCloPisId() != null ? pi.getCloPisId() : "PI")
                            + (pi.getCloPisDescription() != null ? ": " + pi.getCloPisDescription() : "");
                    entries.add(new PiEntry(pi.getCloPisId(), label, indicatorLabel));
                }
            }
        }
        return entries;
    }

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
                        entries.add(new PiEntry(score.getPiId(), label, deriveIndicatorLabel(score.getPiId())));
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

    protected String blockTitle(EvaluationForm form) {
        if (form == null) {
            return "Phi\u1ebfu ch\u1ea5m \u0111i\u1ec3m";
        }
        String title = nullSafe(form.getTitle());
        if (!title.isEmpty()) {
            return title;
        }
        String reportWeek = nullSafe(form.getReportWeek());
        if (!reportWeek.isEmpty()) {
            return "Phi\u1ebfu ch\u1ea5m \u0111i\u1ec3m \u0111\u1ee3t " + reportWeek;
        }
        return "Phi\u1ebfu ch\u1ea5m \u0111i\u1ec3m";
    }

    private String blockSubtitle(EvaluationForm form) {
        return "";
    }

    private String buildIndicatorLabel(Indicator indicator) {
        if (indicator == null) {
            return "CLO";
        }
        String name = nullSafe(indicator.getCloName());
        String description = nullSafe(indicator.getCloDescription());
        if (!name.isEmpty() && !description.isEmpty()) {
            return name + ": " + description;
        }
        return !name.isEmpty() ? name : (!description.isEmpty() ? description : "CLO");
    }

    private String deriveIndicatorLabel(String piId) {
        if (piId == null) {
            return "CLO";
        }
        String trimmed = piId.trim();
        if (trimmed.isEmpty()) {
            return "CLO";
        }
        String upper = trimmed.toUpperCase();
        if (upper.startsWith("C") && upper.length() > 1) {
            String digits = upper.substring(1);
            int dot = digits.indexOf('.');
            if (dot >= 0) {
                digits = digits.substring(0, dot);
            }
            if (!digits.isEmpty()) {
                return "CLO" + digits;
            }
        }
        return "CLO";
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
        if (fullName == null || fullName.trim().isEmpty()) {
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

    private void setIndicatorHeader(Sheet sheet,
                                    Row row,
                                    int start,
                                    int end,
                                    String label,
                                    CellStyle style) {
        if (start > end) {
            return;
        }
        merge(sheet, row.getRowNum(), row.getRowNum(), start, end);
        setCell(row, start, label, style);
    }

    private void applyHorizontalBorder(Sheet sheet, int rowIndex, int startColumn, int endColumn, CellStyle style) {
        for (int col = startColumn; col <= endColumn; col++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }

            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }

            cell.setCellStyle(style);
        }
    }

    private record PiEntry(String piId, String label, String indicatorLabel) {}

    private record FormBlock(EvaluationForm form,
                             List<PiEntry> piEntries,
                             CellStyle headerStyle,
                             CellStyle cellStyle) {}

    private record EvaluationTemplate(String evaluationId,
                                      String title,
                                      List<PiEntry> piEntries) {}

    protected static record ExtraColumn(String header,
                                        int width,
                                        CellStyle headerStyle,
                                        CellStyle cellStyle,
                                        ExtraValueProvider valueProvider) {}

    protected static record SheetLayout(String sheetName,
                                        ColumnType[] columns,
                                        ExtraColumnBuilder extraColumnBuilder) {}

    @FunctionalInterface
    protected interface ExtraValueProvider {
        Object apply(Student student);
    }

    @FunctionalInterface
    protected interface ExtraColumnBuilder {
        List<ExtraColumn> build(Styles styles);
    }

    private enum ColumnType {
        ORDER("STT", 5 * 256),
        STUDENT_ID("M\u00e3 sinh vi\u00ean", 16 * 256),
        STUDENT_LAST("H\u1ecd v\u00e0 t\u00ean \u0111\u1ec7m", 24 * 256),
        STUDENT_FIRST("T\u00ean", 16 * 256),
        GUIDER("Gi\u00e1o vi\u00ean h\u01b0\u1edbng d\u1eabn", 26 * 256),
        PROJECT("T\u00ean \u0111\u1ec1 t\u00e0i \u0111\u1ed3 \u00e1n t\u1ed1t nghi\u1ec7p", 42 * 256),
        REVIEWER("Gi\u1ea3ng vi\u00ean ph\u1ea3n bi\u1ec7n", 26 * 256);

        private final String header;
        private final int width;

        ColumnType(String header, int width) {
            this.header = header;
            this.width = width;
        }

        public String header() {
            return header;
        }

        public int width() {
            return width;
        }
    }

    protected static class Styles {
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













