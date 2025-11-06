package com.javaweb.service;

import com.javaweb.dto.EvaluationDto;
import com.javaweb.dto.EvaluationDto.EvaluationForm;
import com.javaweb.dto.EvaluationDto.Indicator;
import com.javaweb.dto.EvaluationDto.Lecturer;
import com.javaweb.dto.EvaluationDto.Pi;
import com.javaweb.dto.EvaluationDto.Score;
import com.javaweb.dto.EvaluationDto.StudentEvaluation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Service
public class EvaluationExportService {

    /** Build the Excel workbook from the payload. Each lecturer gets a dedicated sheet. */
    public Workbook buildWorkbook(EvaluationDto.Root root) {
        Workbook workbook = new XSSFWorkbook();
        Styles styles = new Styles(workbook);

        EvaluationForm form = root != null ? root.getEvaluationForm() : null;
        List<Lecturer> lecturers = root != null && root.getLecturers() != null
                ? root.getLecturers()
                : Collections.emptyList();
        List<Indicator> indicators = form != null && form.getIndicators() != null
                ? form.getIndicators()
                : Collections.emptyList();

        if (lecturers.isEmpty()) {
            Sheet sheet = workbook.createSheet("Export");
            Row row = sheet.createRow(0);
            setCell(row, 0, "Không có dữ liệu để xuất", styles.normalLeft);
            sheet.autoSizeColumn(0);
            return workbook;
        }

        int sheetIndex = 1;
        for (Lecturer lecturer : lecturers) {
            String name = lecturer.getLecturerName() != null ? lecturer.getLecturerName() : "Giảng viên";
            String sheetName = WorkbookUtil.createSafeSheetName(String.format("%02d-%s", sheetIndex++, name));
            Sheet sheet = workbook.createSheet(sheetName);
            buildLecturerSheet(sheet, form, lecturer, indicators, styles);
        }

        return workbook;
    }

    private void buildLecturerSheet(Sheet sheet,
                                    EvaluationForm form,
                                    Lecturer lecturer,
                                    List<Indicator> indicators,
                                    Styles styles) {

        List<Indicator> indicatorList = indicators != null ? indicators : Collections.emptyList();
        List<List<Pi>> groupedPis = new ArrayList<>();
        int totalPiColumns = 0;

        for (Indicator indicator : indicatorList) {
            List<Pi> pis = indicator != null && indicator.getPis() != null
                    ? indicator.getPis()
                    : Collections.emptyList();
            groupedPis.add(pis);
            totalPiColumns += Math.max(1, pis.size());
        }

        int baseColumns = 4; // TT, Mã SV, Họ và tên, Lớp
        int totalColumns = baseColumns + totalPiColumns + 1; // +1 cho cột Tổng điểm

        configureColumnWidths(sheet, totalColumns);

        int rowIndex = 0;
        rowIndex = buildSheetHeaderBlock(sheet, rowIndex, form, lecturer, totalColumns - 1, styles);
        rowIndex = buildEvaluationTableHeader(sheet, rowIndex, indicatorList, groupedPis, styles, baseColumns);
        populateScores(sheet, rowIndex, lecturer, groupedPis, styles, baseColumns);
    }

    private int buildSheetHeaderBlock(Sheet sheet,
                                      int rowIndex,
                                      EvaluationForm form,
                                      Lecturer lecturer,
                                      int lastColumnIndex,
                                      Styles styles) {

        String academicYear = form != null ? nullSafe(form.getAcademicYear()) : "";
        String formTitle = form != null ? nullSafe(form.getTitle()) : "Phiếu đánh giá";

        int leftBlockEnd = Math.min(4, lastColumnIndex);
        final int minimumRightSpan = 4;
        final int minimumGap = 2;

        int rightBlockStart = -1;
        int available = lastColumnIndex - leftBlockEnd;
        if (available >= minimumRightSpan + minimumGap) {
            int candidate = lastColumnIndex - (minimumRightSpan - 1);
            int minStart = leftBlockEnd + minimumGap;
            if (candidate < minStart) {
                candidate = minStart;
            }
            if (candidate <= lastColumnIndex) {
                rightBlockStart = candidate;
            }
        } else if (available >= minimumRightSpan) {
            rightBlockStart = leftBlockEnd + minimumGap;
        }

        Row row0 = sheet.createRow(rowIndex++);
        if (leftBlockEnd > 0) {
            merge(sheet, row0.getRowNum(), row0.getRowNum(), 0, leftBlockEnd);
        }
        setCell(row0, 0, "Biểu mẫu ĐATN.03A", styles.italicLeft);

        Row row1 = sheet.createRow(rowIndex++);
        if (leftBlockEnd > 0) {
            merge(sheet, row1.getRowNum(), row1.getRowNum(), 0, leftBlockEnd);
        }
        setCell(row1, 0, "BỘ THÔNG TIN VÀ TRUYỀN THÔNG", styles.boldLeft);
        if (rightBlockStart != -1) {
            merge(sheet, row1.getRowNum(), row1.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row1, rightBlockStart, "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM", styles.boldCenter);
        }

        Row row2 = sheet.createRow(rowIndex++);
        if (leftBlockEnd > 0) {
            merge(sheet, row2.getRowNum(), row2.getRowNum(), 0, leftBlockEnd);
        }
        setCell(row2, 0, "HỌC VIỆN CÔNG NGHỆ BƯU CHÍNH VIỄN THÔNG", styles.boldLeft);
        if (rightBlockStart != -1) {
            merge(sheet, row2.getRowNum(), row2.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row2, rightBlockStart, "Độc lập - Tự do - Hạnh phúc", styles.boldUnderlineCenter);
        }

        Row row3 = sheet.createRow(rowIndex++);
        if (rightBlockStart != -1) {
            merge(sheet, row3.getRowNum(), row3.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row3, rightBlockStart, "Hà Nội, ngày .... tháng .... năm ....", styles.centerItalic);
        }

        rowIndex++; // dòng trống

        Row titleRow = sheet.createRow(rowIndex++);
        merge(sheet, titleRow.getRowNum(), titleRow.getRowNum(), 0, lastColumnIndex);
        setCell(titleRow, 0, formTitle.toUpperCase(), styles.title);

        Row subTitle = sheet.createRow(rowIndex++);
        merge(sheet, subTitle.getRowNum(), subTitle.getRowNum(), 0, lastColumnIndex);
        setCell(subTitle, 0, "Đối với Đồ án tốt nghiệp", styles.normalCenter);

        rowIndex++; // dòng trống

        Row section1 = sheet.createRow(rowIndex++);
        setCell(section1, 0, "I. THÔNG TIN CHUNG", styles.boldLeft);

        String lecturerName = nullSafe(lecturer.getLecturerName());
        String lecturerRole = nullSafe(lecturer.getRole());
        rowIndex = writeInfoRow(sheet, rowIndex,
                "Chương trình đào tạo đại học chính quy:", "",
                "Niên khóa:", academicYear,
                styles, lastColumnIndex);

        rowIndex = writeInfoRow(sheet, rowIndex,
                "Hội đồng chuyên môn số:", "",
                null, null,
                styles, lastColumnIndex);

        rowIndex = writeInfoRow(sheet, rowIndex,
                "Họ và tên người chấm ĐATN: " + lecturerName, null,
                "Chức danh trong hội đồng: " + lecturerRole, null,
                styles, lastColumnIndex);

        // Skip optional "Don vi cong tac" row to avoid repeating payload data

        rowIndex++; // dòng trống

        Row section2 = sheet.createRow(rowIndex++);
        setCell(section2, 0, "II. KẾT QUẢ ĐÁNH GIÁ", styles.boldLeft);

        Row note = sheet.createRow(rowIndex++);
        merge(sheet, note.getRowNum(), note.getRowNum(), 0, lastColumnIndex);
        setCell(note, 0, "Điểm mỗi tiêu chí tính theo thang điểm 10, làm tròn đến một chữ số thập phân.", styles.note);
        rowIndex++; // dòng trống giữa chú thích và bảng

        return rowIndex;
    }

    private int buildEvaluationTableHeader(Sheet sheet,
                                           int startRow,
                                           List<Indicator> indicators,
                                           List<List<Pi>> groupedPis,
                                           Styles styles,
                                           int baseColumns) {
        int headerRowIndex = startRow;
        Row row0 = sheet.createRow(headerRowIndex);
        Row row1 = sheet.createRow(headerRowIndex + 1);
        Row row2 = sheet.createRow(headerRowIndex + 2);
        Row row3 = sheet.createRow(headerRowIndex + 3);

        String[] baseHeaders = {"TT", "Mã SV", "Họ và tên SV", "Lớp"};
        for (int i = 0; i < baseHeaders.length; i++) {
            merge(sheet, headerRowIndex, headerRowIndex + 3, i, i);
            setCell(row0, i, baseHeaders[i], styles.header);
            setCell(row1, i, "", styles.header);
            setCell(row2, i, "", styles.header);
            setCell(row3, i, "", styles.header);
        }

        int cloStartCol = baseColumns;
        int columnIndex = baseColumns;
        for (int i = 0; i < indicators.size(); i++) {
            Indicator indicator = indicators.get(i);
            List<Pi> pis = groupedPis.get(i);
            List<Pi> effectivePis = (pis == null || pis.isEmpty())
                    ? Collections.singletonList(null)
                    : pis;

            for (Pi pi : effectivePis) {
                String cloLabel = indicator != null ? nullSafe(indicator.getCloName()) : "";
                String piLabel = pi != null ? nullSafe(pi.getCloPisId()) : "";
                String percent = "";
                if (pi != null && pi.getCloPisWeight() != null) {
                    percent = formatPercent(pi.getCloPisWeight());
                } else if (indicator != null && indicator.getWeight() != null) {
                    percent = formatPercent(indicator.getWeight());
                }

                setCell(row1, columnIndex, cloLabel, styles.header);
                setCell(row2, columnIndex, piLabel, styles.header);
                setCell(row3, columnIndex, percent, styles.headerRed);
                columnIndex++;
            }
        }

        int cloEndCol = columnIndex - 1;
        if (cloEndCol >= cloStartCol) {
            merge(sheet, headerRowIndex, headerRowIndex, cloStartCol, cloEndCol);
            setCell(row0, cloStartCol, "Kết quả đánh giá CLO và tiêu chí", styles.header);
        }

        merge(sheet, headerRowIndex, headerRowIndex + 3, columnIndex, columnIndex);
        setCell(row0, columnIndex, "Tổng điểm", styles.header);
        setCell(row1, columnIndex, "", styles.header);
        setCell(row2, columnIndex, "", styles.header);
        setCell(row3, columnIndex, "", styles.header);

        return headerRowIndex + 4;
    }

    private void populateScores(Sheet sheet,
                                int startRow,
                                Lecturer lecturer,
                                List<List<Pi>> groupedPis,
                                Styles styles,
                                int baseColumns) {

        List<StudentEvaluation> evaluations = lecturer.getEvaluations() != null
                ? lecturer.getEvaluations()
                : Collections.emptyList();

        List<Pi> flattenedPis = new ArrayList<>();
        for (List<Pi> pis : groupedPis) {
            if (pis.isEmpty()) {
                flattenedPis.add(null);
            } else {
                flattenedPis.addAll(pis);
            }
        }

        int lastDataColumn = baseColumns + flattenedPis.size();

        int rowIdx = startRow;
        int order = 1;
        for (StudentEvaluation evaluation : evaluations) {
            Row row = sheet.createRow(rowIdx++);
            setCell(row, 0, order++, styles.cellCenter);
            setCell(row, 1, nullSafe(evaluation.getStudentId()), styles.cellCenter);
            setCell(row, 2, "", styles.cellLeft);
            setCell(row, 3, nullSafe(evaluation.getClassName()), styles.cellCenter);

            Map<String, Double> scores = toScoreMap(evaluation);
            int colIdx = baseColumns;
            for (Pi pi : flattenedPis) {
                Double value = null;
                if (pi != null && pi.getCloPisId() != null) {
                    value = scores.get(pi.getCloPisId());
                }
                setCell(row, colIdx++, value, styles.cellCenter);
            }

            Double total = evaluation.getEvaluations() != null
                    ? evaluation.getEvaluations().getTotalScore()
                    : null;
            setCell(row, colIdx, total, styles.cellCenter);
        }

        int noteRow = rowIdx + 1;

        noteRow++; // spacer

        String[] notes = {
                "Ghi ch\u00fa:",
                "CLO 3 Thi\u1ebft k\u1ebf ph\u1ea7n c\u1ee9ng v\u00e0 ph\u1ea7n m\u1ec1m, ph\u00e2n t\u00edch d\u1eef li\u1ec7u \u0111\u1ec3 \u0111\u00e1nh gi\u00e1 hi\u1ec7u qu\u1ea3 ho\u1ea1t \u0111\u1ed9ng c\u1ee7a h\u1ec7 th\u1ed1ng \u0111i\u1ec7n t\u1eed \u0111i\u1ec7n t\u1eed.",
                "C3.3 Ti\u1ebfn h\u00e0nh \u0111\u01b0\u1ee3c c\u00e1c th\u00ed nghi\u1ec7m, c\u0169ng nh\u01b0 ph\u00e2n t\u00edch, \u0111\u00e1nh gi\u00e1 v\u00e0 di\u1ec5n gi\u1ea3i c\u00e1c k\u1ebft qu\u1ea3 th\u00ed nghi\u1ec7m.",
                "C4.0 Th\u1ec3 hi\u1ec7n \u0111\u01b0\u1ee3c \u0111\u1ea1o \u0111\u1ee9c v\u00e0 tr\u00e1ch nhi\u1ec7m ngh\u1ec1 nghi\u1ec7p trong qu\u00e1 tr\u00ecnh tri\u1ec3n khai c\u00e1c h\u1ec7 th\u1ed1ng \u0111i\u1ec7n.",
                "C4.2 Gi\u1ea3i th\u00edch \u0111\u01b0\u1ee3c t\u00e1c \u0111\u1ed9ng c\u1ee7a k\u1ebft qu\u1ea3 nghi\u00ean c\u1ee9u \u0111\u1ed1i v\u1edbi c\u1ed9ng \u0111\u1ed3ng, x\u00e3 h\u1ed9i, ho\u1eb7c ng\u00e0nh ngh\u1ec1.",
                "C5.3 Hi\u1ec7u qu\u1ea3 gi\u1ea3i quy\u1ebft v\u1ea5n \u0111\u1ec1 c\u1ee7a nh\u00f3m.",
                "CLO 6 V\u1eadn d\u1ee5ng k\u1ef9 n\u0103ng giao ti\u1ebfp trong ng\u00e0nh \u0111i\u1ec7n - \u0111i\u1ec7n t\u1eed.",
                "C6.3 Kh\u1ea3 n\u0103ng thuy\u1ebft tr\u00ecnh.",
                "C6.4 Kh\u1ea3 n\u0103ng giao ti\u1ebfp \u0111\u1ed1i tho\u1ea1i v\u00e0 tr\u1ea3 l\u1eddi c\u00e1c c\u00e2u h\u1ecfi c\u1ee7a h\u1ed9i \u0111\u1ed3ng."
        };

        for (int i = 0; i < notes.length; i++) {
            Row noteLine = sheet.createRow(noteRow++);
            merge(sheet, noteLine.getRowNum(), noteLine.getRowNum(), 0, lastDataColumn);
            CellStyle noteStyle = (i == 0) ? styles.noteHeading : styles.noteEmphasis;
            setCell(noteLine, 0, notes[i], noteStyle);
        }

        noteRow++; // blank line before signature block

        int signatureStartColumn = Math.min(1, lastDataColumn);
        int signatureEndColumn = Math.max(signatureStartColumn, Math.min(signatureStartColumn + 3, lastDataColumn));

        Row signerTitle = sheet.createRow(noteRow++);
        merge(sheet, signerTitle.getRowNum(), signerTitle.getRowNum(), signatureStartColumn, signatureEndColumn);
        setCell(signerTitle, signatureStartColumn, "NG\u01af\u1edcI \u0110\u00c1NH GI\u00c1", styles.boldLeft);

        Row signerSpace = sheet.createRow(noteRow++);
        merge(sheet, signerSpace.getRowNum(), signerSpace.getRowNum(), signatureStartColumn, signatureEndColumn);
        setCell(signerSpace, signatureStartColumn, "", styles.normalLeft);

        Row signerSpace2 = sheet.createRow(noteRow++);
        merge(sheet, signerSpace2.getRowNum(), signerSpace2.getRowNum(), signatureStartColumn, signatureEndColumn);
        setCell(signerSpace2, signatureStartColumn, "", styles.normalLeft);

        Row signerName = sheet.createRow(noteRow);
        merge(sheet, signerName.getRowNum(), signerName.getRowNum(), signatureStartColumn, signatureEndColumn);
        setCell(signerName, signatureStartColumn, nullSafe(lecturer.getLecturerName()).toUpperCase(), styles.boldLeft);
    }

    private int writeInfoRow(Sheet sheet,
                             int rowIndex,
                             String leftLabel,
                             String leftValue,
                             String rightLabel,
                             String rightValue,
                             Styles styles,
                             int lastColumnIndex) {

        Row row = sheet.createRow(rowIndex);
        int nextColumn = 0;

        if (leftLabel != null && !leftLabel.isEmpty()) {
            int leftLabelEnd = Math.min(4, lastColumnIndex);
            if (leftLabelEnd > 0) {
                merge(sheet, rowIndex, rowIndex, 0, leftLabelEnd);
            }
            setCell(row, 0, leftLabel, styles.normalLeft);
            nextColumn = leftLabelEnd + 1;

            if (leftValue != null && !leftValue.isEmpty() && nextColumn <= lastColumnIndex) {
                int valueEnd = Math.min(nextColumn + 2, lastColumnIndex);
                merge(sheet, rowIndex, rowIndex, nextColumn, valueEnd);
                setCell(row, nextColumn, leftValue, styles.boldLeft);
                nextColumn = valueEnd + 1;
            }
        }

        if (rightLabel != null && !rightLabel.isEmpty() && nextColumn <= lastColumnIndex) {
            String safeValue = rightValue != null ? rightValue : "";
            if (safeValue.isEmpty()) {
                int end = lastColumnIndex;
                merge(sheet, rowIndex, rowIndex, nextColumn, end);
                setCell(row, nextColumn, rightLabel, styles.normalLeft);
            } else {
                int labelEnd = Math.min(nextColumn + 1, lastColumnIndex);
                merge(sheet, rowIndex, rowIndex, nextColumn, labelEnd);
                setCell(row, nextColumn, rightLabel, styles.normalLeft);
                int valueStart = Math.min(labelEnd + 1, lastColumnIndex);
                if (valueStart <= lastColumnIndex) {
                    int valueEnd = lastColumnIndex;
                    merge(sheet, rowIndex, rowIndex, valueStart, valueEnd);
                    setCell(row, valueStart, safeValue, styles.boldLeft);
                }
            }
        }

        return rowIndex + 1;
    }

    private void configureColumnWidths(Sheet sheet, int totalColumns) {
        sheet.setColumnWidth(0, 6 * 256);
        sheet.setColumnWidth(1, 18 * 256);
        sheet.setColumnWidth(2, 28 * 256);
        sheet.setColumnWidth(3, 14 * 256);
        for (int i = 4; i < totalColumns - 1; i++) {
            sheet.setColumnWidth(i, 12 * 256);
        }
        sheet.setColumnWidth(Math.max(4, totalColumns - 1), 14 * 256);
    }

    private Map<String, Double> toScoreMap(StudentEvaluation evaluation) {
        Map<String, Double> map = new HashMap<>();
        if (evaluation.getEvaluations() == null || evaluation.getEvaluations().getScores() == null) {
            return map;
        }
        for (Score score : evaluation.getEvaluations().getScores()) {
            if (score.getPiId() != null && score.getScore() != null) {
                map.putIfAbsent(score.getPiId(), score.getScore());
            }
        }
        return map;
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

    private String nullSafe(String value) {
        return value != null ? value : "";
    }

    private String formatPercent(Double value) {
        return value != null ? String.format("%.0f%%", value * 100) : "";
    }

    private static class Styles {
        final CellStyle boldLeft;
        final CellStyle boldCenter;
        final CellStyle boldUnderlineCenter;
        final CellStyle normalLeft;
        final CellStyle normalCenter;
        final CellStyle centerItalic;
        final CellStyle italicLeft;
        final CellStyle title;
        final CellStyle header;
        final CellStyle headerRed;
        final CellStyle cellCenter;
        final CellStyle cellLeft;
        final CellStyle note;
        final CellStyle noteHeading;
        final CellStyle noteEmphasis;

        Styles(Workbook workbook) {
            Font normal = workbook.createFont();
            normal.setFontName("Times New Roman");
            normal.setFontHeightInPoints((short) 12);

            Font bold = workbook.createFont();
            bold.setBold(true);
            bold.setFontName("Times New Roman");
            bold.setFontHeightInPoints((short) 12);

            Font boldUnderline = workbook.createFont();
            boldUnderline.setBold(true);
            boldUnderline.setUnderline(Font.U_SINGLE);
            boldUnderline.setFontName("Times New Roman");
            boldUnderline.setFontHeightInPoints((short) 12);

            Font redBold = workbook.createFont();
            redBold.setBold(true);
            redBold.setFontName("Times New Roman");
            redBold.setFontHeightInPoints((short) 12);
            redBold.setColor(IndexedColors.RED.getIndex());

            Font titleFont = workbook.createFont();
            titleFont.setBold(true);
            titleFont.setFontName("Times New Roman");
            titleFont.setFontHeightInPoints((short) 14);

            Font italic = workbook.createFont();
            italic.setItalic(true);
            italic.setFontName("Times New Roman");
            italic.setFontHeightInPoints((short) 12);

            Font noteFont = workbook.createFont();
            noteFont.setItalic(true);
            noteFont.setFontName("Times New Roman");
            noteFont.setFontHeightInPoints((short) 12);
            noteFont.setColor(IndexedColors.BLACK.getIndex());

            Font noteHeadingFont = workbook.createFont();
            noteHeadingFont.setBold(true);
            noteHeadingFont.setItalic(true);
            noteHeadingFont.setFontName("Times New Roman");
            noteHeadingFont.setFontHeightInPoints((short) 12);
            noteHeadingFont.setColor(IndexedColors.BLACK.getIndex());

            Font noteEmphasisFont = workbook.createFont();
            noteEmphasisFont.setBold(true);
            noteEmphasisFont.setItalic(true);
            noteEmphasisFont.setFontName("Times New Roman");
            noteEmphasisFont.setFontHeightInPoints((short) 12);
            noteEmphasisFont.setColor(IndexedColors.BLACK.getIndex());

            boldLeft = workbook.createCellStyle();
            boldLeft.setFont(bold);
            boldLeft.setAlignment(HorizontalAlignment.LEFT);
            boldLeft.setVerticalAlignment(VerticalAlignment.CENTER);

            boldCenter = workbook.createCellStyle();
            boldCenter.setFont(bold);
            boldCenter.setAlignment(HorizontalAlignment.CENTER);
            boldCenter.setVerticalAlignment(VerticalAlignment.CENTER);

            boldUnderlineCenter = workbook.createCellStyle();
            boldUnderlineCenter.setFont(boldUnderline);
            boldUnderlineCenter.setAlignment(HorizontalAlignment.CENTER);
            boldUnderlineCenter.setVerticalAlignment(VerticalAlignment.CENTER);

            normalLeft = workbook.createCellStyle();
            normalLeft.setFont(normal);
            normalLeft.setAlignment(HorizontalAlignment.LEFT);
            normalLeft.setVerticalAlignment(VerticalAlignment.CENTER);

            normalCenter = workbook.createCellStyle();
            normalCenter.setFont(normal);
            normalCenter.setAlignment(HorizontalAlignment.CENTER);
            normalCenter.setVerticalAlignment(VerticalAlignment.CENTER);

            centerItalic = workbook.createCellStyle();
            centerItalic.setFont(italic);
            centerItalic.setAlignment(HorizontalAlignment.CENTER);
            centerItalic.setVerticalAlignment(VerticalAlignment.CENTER);

            italicLeft = workbook.createCellStyle();
            italicLeft.setFont(italic);
            italicLeft.setAlignment(HorizontalAlignment.LEFT);
            italicLeft.setVerticalAlignment(VerticalAlignment.CENTER);

            title = workbook.createCellStyle();
            title.setFont(titleFont);
            title.setAlignment(HorizontalAlignment.CENTER);
            title.setVerticalAlignment(VerticalAlignment.CENTER);

            header = workbook.createCellStyle();
            header.setFont(bold);
            header.setAlignment(HorizontalAlignment.CENTER);
            header.setVerticalAlignment(VerticalAlignment.CENTER);
            header.setWrapText(true);
            addBorder(header);

            headerRed = workbook.createCellStyle();
            headerRed.setFont(redBold);
            headerRed.setAlignment(HorizontalAlignment.CENTER);
            headerRed.setVerticalAlignment(VerticalAlignment.CENTER);
            headerRed.setWrapText(true);
            addBorder(headerRed);

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

            note = workbook.createCellStyle();
            note.setFont(noteFont);
            note.setAlignment(HorizontalAlignment.LEFT);
            note.setVerticalAlignment(VerticalAlignment.CENTER);

            noteHeading = workbook.createCellStyle();
            noteHeading.setFont(noteHeadingFont);
            noteHeading.setAlignment(HorizontalAlignment.LEFT);
            noteHeading.setVerticalAlignment(VerticalAlignment.CENTER);
            noteHeading.setWrapText(true);

            noteEmphasis = workbook.createCellStyle();
            noteEmphasis.setFont(noteEmphasisFont);
            noteEmphasis.setAlignment(HorizontalAlignment.LEFT);
            noteEmphasis.setVerticalAlignment(VerticalAlignment.CENTER);
            noteEmphasis.setWrapText(true);
        }

        private void addBorder(CellStyle style) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
        }
    }
}
