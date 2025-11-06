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

        Row signature = sheet.createRow(rowIdx + 1);
        setCell(signature, 0, "Chữ ký và họ tên:", styles.normalLeft);
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
            noteFont.setColor(IndexedColors.RED.getIndex());

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
        }

        private void addBorder(CellStyle style) {
            style.setBorderTop(BorderStyle.THIN);
            style.setBorderBottom(BorderStyle.THIN);
            style.setBorderLeft(BorderStyle.THIN);
            style.setBorderRight(BorderStyle.THIN);
        }
    }
}
