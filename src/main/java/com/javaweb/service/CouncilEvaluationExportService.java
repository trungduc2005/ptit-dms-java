package com.javaweb.service;

import com.javaweb.dto.CouncilEvaluationDto;
import com.javaweb.dto.CouncilEvaluationDto.EvaluationForm;
import com.javaweb.dto.CouncilEvaluationDto.Indicator;
import com.javaweb.dto.CouncilEvaluationDto.Lecturer;
import com.javaweb.dto.CouncilEvaluationDto.Pi;
import com.javaweb.dto.CouncilEvaluationDto.Score;
import com.javaweb.dto.CouncilEvaluationDto.StudentEvaluation;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

@Service
public class CouncilEvaluationExportService {

    /** Build the Excel workbook from the payload. Each lecturer gets a dedicated sheet. */
    public Workbook buildWorkbook(CouncilEvaluationDto.Root root) {
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
            setCell(row, 0, "Kh\u00f4ng c\u00f3 d\u1eef li\u1ec7u \u0111\u1ec3 xu\u1ea5t", styles.normalLeft);
            sheet.autoSizeColumn(0);
            return workbook;
        }

        int sheetIndex = 1;
        for (Lecturer lecturer : lecturers) {
            String name = lecturer.getLecturerName() != null ? lecturer.getLecturerName() : "Gi\u1ea3ng vi\u00ean";
            String sheetName = WorkbookUtil.createSafeSheetName(String.format("%02d-%s", sheetIndex++, name));
            Sheet sheet = workbook.createSheet(sheetName);
            buildLecturerSheet(sheet, form, lecturer, indicators, styles);
        }

        buildSummarySheet(workbook, form, lecturers, styles);

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

        int baseColumns = 5; // TT, M\u00e3 SV, H\u1ecd, T\u00ean, L\u1edbp
        int totalColumns = baseColumns + totalPiColumns + 2; // +1 T\u1ed5ng \u0111i\u1ec3m, +1 Nh\u1eadn x\u00e9t

        configureColumnWidths(sheet, totalColumns);

        int rowIndex = 0;
        rowIndex = buildSheetHeaderBlock(sheet, rowIndex, form, lecturer, totalColumns - 1, styles);
        rowIndex = buildEvaluationTableHeader(sheet, rowIndex, indicatorList, groupedPis, styles, baseColumns);
        populateScores(sheet, rowIndex, lecturer, groupedPis, styles, baseColumns);
    }

    private void buildSummarySheet(Workbook workbook,
                                   EvaluationForm form,
                                   List<Lecturer> lecturers,
                                   Styles styles) {
        if (lecturers == null || lecturers.isEmpty()) {
            return;
        }

        String summaryName = WorkbookUtil.createSafeSheetName("00-Tong hop");
        Sheet sheet = workbook.createSheet(summaryName);
        configureSummaryColumnWidths(sheet, lecturers.size());

        int rowIndex = 0;
        rowIndex = buildSummaryHeader(sheet, rowIndex, lecturers, styles);

        List<SummaryEntry> summaries = buildSummaryEntries(lecturers);
        populateSummaryRows(sheet, rowIndex, summaries, styles);
    }

    private void configureSummaryColumnWidths(Sheet sheet, int lecturerCount) {
        sheet.setColumnWidth(0, 6 * 256);
        sheet.setColumnWidth(1, 18 * 256);
        sheet.setColumnWidth(2, 22 * 256);
        sheet.setColumnWidth(3, 14 * 256);
        sheet.setColumnWidth(4, 18 * 256);
        int scoreStart = 5;
        for (int i = 0; i < lecturerCount; i++) {
            sheet.setColumnWidth(scoreStart + i, 12 * 256);
        }
        sheet.setColumnWidth(scoreStart + lecturerCount, 14 * 256);
        sheet.setColumnWidth(scoreStart + lecturerCount + 1, 28 * 256);
    }

    private int buildSummaryHeader(Sheet sheet,
                                   int startRow,
                                   List<Lecturer> lecturers,
                                   Styles styles) {
        int lecturerCount = lecturers.size();
        Row row0 = sheet.createRow(startRow);
        Row row1 = sheet.createRow(startRow + 1);
        Row row2 = sheet.createRow(startRow + 2);
        Row row3 = sheet.createRow(startRow + 3);

        merge(sheet, startRow, startRow + 3, 0, 0);
        setCell(row0, 0, "TT", styles.header);
        setCell(row1, 0, "", styles.header);
        setCell(row2, 0, "", styles.header);
        setCell(row3, 0, "", styles.header);

        merge(sheet, startRow, startRow + 3, 1, 1);
        setCell(row0, 1, "M\u00e3 SV", styles.header);
        setCell(row1, 1, "", styles.header);
        setCell(row2, 1, "", styles.header);
        setCell(row3, 1, "", styles.header);

        CellRangeAddress nameRegion = new CellRangeAddress(startRow, startRow + 3, 2, 3);
        merge(sheet, nameRegion.getFirstRow(), nameRegion.getLastRow(), nameRegion.getFirstColumn(), nameRegion.getLastColumn());
        applyHeaderBorder(sheet, nameRegion);
        setCell(row0, 2, "H\u1ecd v\u00e0 t\u00ean SV", styles.header);
        setCell(row0, 3, "", styles.header);
        setCell(row1, 2, "", styles.header);
        setCell(row1, 3, "", styles.header);
        setCell(row2, 2, "", styles.header);
        setCell(row2, 3, "", styles.header);
        setCell(row3, 2, "", styles.header);
        setCell(row3, 3, "", styles.header);

        merge(sheet, startRow, startRow + 3, 4, 4);
        setCell(row0, 4, "L\u1edbp", styles.header);
        setCell(row1, 4, "", styles.header);
        setCell(row2, 4, "", styles.header);
        setCell(row3, 4, "", styles.header);

        int scoreStart = 5;
        int scoreEnd = scoreStart + lecturerCount - 1;
        if (lecturerCount > 0) {
            CellRangeAddress summaryRegion = new CellRangeAddress(startRow, startRow, scoreStart, scoreEnd);
            merge(sheet, summaryRegion.getFirstRow(), summaryRegion.getLastRow(), summaryRegion.getFirstColumn(), summaryRegion.getLastColumn());
            applyHeaderBorder(sheet, summaryRegion);
            setCell(row0, scoreStart, "GPA t\u1ed5ng k\u1ebft", styles.header);
            for (int i = 0; i < lecturerCount; i++) {
                Lecturer lecturer = lecturers.get(i);
                String lecturerName = lecturer != null ? nullSafe(lecturer.getLecturerName()) : "";
                int columnIndex = scoreStart + i;
                setCell(row1, columnIndex, lecturerName, styles.header);
                setCell(row2, columnIndex, String.format("GPA%d", i + 1), styles.header);
                setCell(row3, columnIndex, "", styles.header);
            }
        }

        int gpaColumn = scoreStart + lecturerCount;
        CellRangeAddress gpaRegion = new CellRangeAddress(startRow, startRow + 3, gpaColumn, gpaColumn);
        merge(sheet, gpaRegion.getFirstRow(), gpaRegion.getLastRow(), gpaRegion.getFirstColumn(), gpaRegion.getLastColumn());
        applyHeaderBorder(sheet, gpaRegion);
        setCell(row0, gpaColumn, "\u0110i\u1ec3m GPA (%)", styles.summaryGpaHeader);
        setCell(row1, gpaColumn, "", styles.summaryGpaHeader);
        setCell(row2, gpaColumn, "", styles.summaryGpaHeader);
        setCell(row3, gpaColumn, "", styles.summaryGpaHeader);

        int commentColumn = gpaColumn + 1;
        CellRangeAddress commentRegion = new CellRangeAddress(startRow, startRow + 3, commentColumn, commentColumn);
        merge(sheet, commentRegion.getFirstRow(), commentRegion.getLastRow(), commentRegion.getFirstColumn(), commentRegion.getLastColumn());
        applyHeaderBorder(sheet, commentRegion);
        setCell(row0, commentColumn, "Nh\u1eadn x\u00e9t kh\u00e1c \u0111\u1ed1i v\u1edbi sinh vi\u00ean", styles.header);
        setCell(row1, commentColumn, "", styles.header);
        setCell(row2, commentColumn, "", styles.header);
        setCell(row3, commentColumn, "", styles.header);

        return startRow + 4;
    }

    private List<SummaryEntry> buildSummaryEntries(List<Lecturer> lecturers) {
        int lecturerCount = lecturers.size();
        Map<String, SummaryEntry> map = new LinkedHashMap<>();

        for (int lecturerIndex = 0; lecturerIndex < lecturerCount; lecturerIndex++) {
            Lecturer lecturer = lecturers.get(lecturerIndex);
            List<StudentEvaluation> evaluations = lecturer != null && lecturer.getEvaluations() != null
                    ? lecturer.getEvaluations()
                    : Collections.emptyList();
            String lecturerName = lecturer != null ? nullSafe(lecturer.getLecturerName()) : "";

            for (StudentEvaluation evaluation : evaluations) {
                if (evaluation == null) {
                    continue;
                }
                String studentId = nullSafe(evaluation.getStudentId());
                String className = nullSafe(evaluation.getClassName());
                String studentName = nullSafe(evaluation.getStudentName());
                String key = studentId + "|" + className + "|" + studentName;
                SummaryEntry entry = map.computeIfAbsent(key,
                        k -> new SummaryEntry(studentId, studentName, className, lecturerCount));

                if (entry.studentName == null || entry.studentName.isEmpty()) {
                    entry.studentName = studentName;
                }
                if (entry.className == null || entry.className.isEmpty()) {
                    entry.className = className;
                }

                Double totalScore = evaluation.getEvaluations() != null
                        ? evaluation.getEvaluations().getTotalScore()
                        : null;
                entry.scores.set(lecturerIndex, totalScore);

                String labeledComment = labelComment(lecturerName, evaluation.getComment());
                if (!labeledComment.isEmpty()) {
                    entry.comments.add(labeledComment);
                }
            }
        }

        return new ArrayList<>(map.values());
    }

    private void populateSummaryRows(Sheet sheet,
                                     int startRow,
                                     List<SummaryEntry> entries,
                                     Styles styles) {
        int rowIndex = startRow;
        int order = 1;
        for (SummaryEntry entry : entries) {
            Row row = sheet.createRow(rowIndex++);
            setCell(row, 0, order++, styles.cellCenter);
            setCell(row, 1, nullSafe(entry.studentId), styles.cellCenter);

            String[] nameParts = splitStudentName(entry.studentName);
            setCell(row, 2, nameParts[0], styles.cellLeft);
            setCell(row, 3, nameParts[1], styles.cellLeft);
            setCell(row, 4, nullSafe(entry.className), styles.cellCenter);

            int colIdx = 5;
            for (Double score : entry.scores) {
                setCell(row, colIdx++, score, styles.cellCenter);
            }

            Double average = averageScore(entry.scores);
            setCell(row, colIdx++, average, styles.summaryGpaCell);

            String comments = String.join("\n", entry.comments).trim();
            CellStyle commentStyle = comments.contains("\n") ? styles.cellLeftWrap : styles.cellLeft;
            setCell(row, colIdx, comments, commentStyle);
        }
    }

    private Double averageScore(List<Double> scores) {
        double sum = 0;
        int count = 0;
        for (Double score : scores) {
            if (score != null) {
                sum += score;
                count++;
            }
        }
        return count == 0 ? null : sum / count;
    }

    private String labelComment(String lecturerName, String rawComment) {
        String comment = rawComment != null ? rawComment.trim() : "";
        if (comment.isEmpty()) {
            return "";
        }
        String name = lecturerName != null ? lecturerName.trim() : "";
        return name.isEmpty() ? comment : name + ": " + comment;
    }

    private int buildSheetHeaderBlock(Sheet sheet,
                                      int rowIndex,
                                      EvaluationForm form,
                                      Lecturer lecturer,
                                      int lastColumnIndex,
                                      Styles styles) {

        String academicYear = form != null ? nullSafe(form.getAcademicYear()) : "";
        String formTitle = form != null ? nullSafe(form.getTitle()) : "Phi\u1ebfu \u0111\u00e1nh gi\u00e1";

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
        setCell(row0, 0, "Bi\u1ec3u m\u1eabu ATN.03A", styles.italicLeft);

        Row row1 = sheet.createRow(rowIndex++);
        if (leftBlockEnd > 0) {
            merge(sheet, row1.getRowNum(), row1.getRowNum(), 0, leftBlockEnd);
        }
        setCell(row1, 0, "B\u1ed8 TH\u00d4NG TIN V\u00c0 TRUY\u1ec0N TH\u00d4NG", styles.boldLeft);
        if (rightBlockStart != -1) {
            merge(sheet, row1.getRowNum(), row1.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row1, rightBlockStart, "C\u1ed8NG H\u00d2A X\u00c3 H\u1ed8I CH\u1ee6 NGH\u0128A VI\u1ec6T NAM", styles.boldCenter);
        }

        Row row2 = sheet.createRow(rowIndex++);
        if (leftBlockEnd > 0) {
            merge(sheet, row2.getRowNum(), row2.getRowNum(), 0, leftBlockEnd);
        }
        setCell(row2, 0, "H\u1eccC VI\u1ec6N C\u00d4NG NGH\u1ec6 B\u01afU CH\u00cdNH VI\u1ec4N TH\u00d4NG", styles.boldLeft);
        if (rightBlockStart != -1) {
            merge(sheet, row2.getRowNum(), row2.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row2, rightBlockStart, "\u0110\u1ed9c l\u1eadp - T\u1ef1 do - H\u1ea1nh ph\u00fac", styles.boldUnderlineCenter);
        }

        Row row3 = sheet.createRow(rowIndex++);
        if (rightBlockStart != -1) {
            merge(sheet, row3.getRowNum(), row3.getRowNum(), rightBlockStart, lastColumnIndex);
            setCell(row3, rightBlockStart, "H\u00e0 N\u1ed9i, ng\u00e0y .... th\u00e1ng .... n\u0103m ....", styles.centerItalic);
        }

        rowIndex++; // d\u00f2ng tr\u1ed1ng

        Row titleRow = sheet.createRow(rowIndex++);
        merge(sheet, titleRow.getRowNum(), titleRow.getRowNum(), 0, lastColumnIndex);
        setCell(titleRow, 0, formTitle.toUpperCase(), styles.title);

        Row subTitle = sheet.createRow(rowIndex++);
        merge(sheet, subTitle.getRowNum(), subTitle.getRowNum(), 0, lastColumnIndex);
        setCell(subTitle, 0, "\u0110\u1ed1i v\u1edbi \u0111\u1ed3 \u00e1n t\u1ed1t nghi\u1ec7p", styles.normalCenter);

        rowIndex++; // d\u00f2ng tr\u1ed1ng

        Row section1 = sheet.createRow(rowIndex++);
        setCell(section1, 0, "I. TH\u00d4NG TIN CHUNG", styles.boldLeft);

        String lecturerName = lecturer != null ? nullSafe(lecturer.getLecturerName()) : "";
        String lecturerRole = lecturer != null ? nullSafe(lecturer.getRole()) : "";
        String lecturerDepartment = lecturer != null ? nullSafe(lecturer.getDepartment()) : "";

        rowIndex = writeInfoRow(sheet, rowIndex,
                "Ch\u01b0\u01a1ng tr\u00ecnh \u0111\u00e0o t\u1ea1o \u0111\u1ea1i h\u1ecdc ch\u00ednh quy:", "",
                "Ni\u00ean kh\u00f3a:", academicYear,
                styles, lastColumnIndex);

        rowIndex = writeInfoRow(sheet, rowIndex,
                "H\u1ed9i \u0111\u1ed3ng chuy\u00ean m\u00f4n s\u1ed1:", "",
                null, null,
                styles, lastColumnIndex);

        rowIndex = writeInfoRow(sheet, rowIndex,
                "H\u1ecd v\u00e0 t\u00ean ng\u01b0\u1eddi ch\u1ea5m \u0110ATN: " + lecturerName, null,
                "Ch\u1ee9c danh trong h\u1ed9i \u0111\u1ed3ng: " + lecturerRole, null,
                styles, lastColumnIndex);

        Row unitRow = sheet.createRow(rowIndex++);
        merge(sheet, unitRow.getRowNum(), unitRow.getRowNum(), 0, lastColumnIndex);
        setCell(unitRow, 0, "\u0110\u01a1n v\u1ecb c\u00f4ng t\u00e1c: " + lecturerDepartment, styles.normalLeft);

        rowIndex++; // d\u00f2ng tr\u1ed1ng

        Row section2 = sheet.createRow(rowIndex++);
        setCell(section2, 0, "II. K\u1ebeT QU\u1ea2 \u0110\u00c1NH GI\u00c1", styles.boldLeft);

        Row note = sheet.createRow(rowIndex++);
        merge(sheet, note.getRowNum(), note.getRowNum(), 0, lastColumnIndex);
        setCell(note, 0, "\u0110i\u1ec3m m\u1ed7i ti\u00eau ch\u00ed t\u00ednh theo thang \u0111i\u1ec3m 10, l\u00e0m tr\u00f2n \u0111\u1ebfn m\u1ed9t ch\u1eef s\u1ed1 th\u1eadp ph\u00e2n.", styles.note);
        rowIndex++; // d\u00f2ng tr\u1ed1ng gi\u1eefa ch\u00fa th\u00edch v\u00e0 b\u1ea3ng

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

        merge(sheet, headerRowIndex, headerRowIndex + 3, 0, 0);
        setCell(row0, 0, "TT", styles.header);
        setCell(row1, 0, "", styles.header);
        setCell(row2, 0, "", styles.header);
        setCell(row3, 0, "", styles.header);

        merge(sheet, headerRowIndex, headerRowIndex + 3, 1, 1);
        setCell(row0, 1, "M\u00e3 SV", styles.header);
        setCell(row1, 1, "", styles.header);
        setCell(row2, 1, "", styles.header);
        setCell(row3, 1, "", styles.header);

        CellRangeAddress nameRegion = new CellRangeAddress(headerRowIndex, headerRowIndex + 3, 2, 3);
        merge(sheet, nameRegion.getFirstRow(), nameRegion.getLastRow(), nameRegion.getFirstColumn(), nameRegion.getLastColumn());
        applyHeaderBorder(sheet, nameRegion);
        setCell(row0, 2, "H\u1ecd v\u00e0 t\u00ean SV", styles.header);
        setCell(row0, 3, "", styles.header);
        setCell(row1, 2, "", styles.header);
        setCell(row1, 3, "", styles.header);
        setCell(row2, 2, "", styles.header);
        setCell(row2, 3, "", styles.header);
        setCell(row3, 2, "", styles.header);
        setCell(row3, 3, "", styles.header);

        merge(sheet, headerRowIndex, headerRowIndex + 3, 4, 4);
        setCell(row0, 4, "L\u1edbp", styles.header);
        setCell(row1, 4, "", styles.header);
        setCell(row2, 4, "", styles.header);
        setCell(row3, 4, "", styles.header);

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
            CellRangeAddress cloRegion = new CellRangeAddress(headerRowIndex, headerRowIndex, cloStartCol, cloEndCol);
            merge(sheet, cloRegion.getFirstRow(), cloRegion.getLastRow(), cloRegion.getFirstColumn(), cloRegion.getLastColumn());
            applyHeaderBorder(sheet, cloRegion);
            setCell(row0, cloStartCol, "K\u1ebft qu\u1ea3 \u0111\u00e1nh gi\u00e1 CLO v\u00e0 ti\u00eau ch\u00ed", styles.header);
        }

        CellRangeAddress totalRegion = new CellRangeAddress(headerRowIndex, headerRowIndex + 3, columnIndex, columnIndex);
        merge(sheet, totalRegion.getFirstRow(), totalRegion.getLastRow(), totalRegion.getFirstColumn(), totalRegion.getLastColumn());
        applyHeaderBorder(sheet, totalRegion);
        setCell(row0, columnIndex, "T\u1ed5ng \u0111i\u1ec3m", styles.header);
        setCell(row1, columnIndex, "", styles.header);
        setCell(row2, columnIndex, "", styles.header);
        setCell(row3, columnIndex, "", styles.header);

        columnIndex++;
        CellRangeAddress commentRegion = new CellRangeAddress(headerRowIndex, headerRowIndex + 3, columnIndex, columnIndex);
        merge(sheet, commentRegion.getFirstRow(), commentRegion.getLastRow(), commentRegion.getFirstColumn(), commentRegion.getLastColumn());
        applyHeaderBorder(sheet, commentRegion);
        setCell(row0, columnIndex, "Nh\u1eadn x\u00e9t kh\u00e1c \u0111\u1ed1i v\u1edbi sinh vi\u00ean", styles.header);
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

        int lastDataColumn = baseColumns + flattenedPis.size() + 1;

        int rowIdx = startRow;
        int order = 1;
        for (StudentEvaluation evaluation : evaluations) {
            Row row = sheet.createRow(rowIdx++);
            setCell(row, 0, order++, styles.cellCenter);
            setCell(row, 1, nullSafe(evaluation.getStudentId()), styles.cellCenter);
            String[] nameParts = splitStudentName(evaluation.getStudentName());
            setCell(row, 2, nameParts[0], styles.cellLeft);
            setCell(row, 3, nameParts[1], styles.cellLeft);
            setCell(row, 4, nullSafe(evaluation.getClassName()), styles.cellCenter);

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
            setCell(row, colIdx++, total, styles.cellCenter);
            setCell(row, colIdx, nullSafe(evaluation.getComment()), styles.cellLeft);
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
        sheet.setColumnWidth(2, 22 * 256);
        sheet.setColumnWidth(3, 14 * 256);
        sheet.setColumnWidth(4, 18 * 256);
        int totalColumnIndex = Math.max(5, totalColumns - 2);
        int commentColumnIndex = Math.max(totalColumnIndex + 1, totalColumns - 1);
        for (int i = 5; i < Math.min(totalColumnIndex, totalColumns); i++) {
            sheet.setColumnWidth(i, 12 * 256);
        }
        if (totalColumnIndex < totalColumns) {
            sheet.setColumnWidth(totalColumnIndex, 14 * 256);
        }
        if (commentColumnIndex < totalColumns) {
            sheet.setColumnWidth(commentColumnIndex, 28 * 256);
        }
    }

    private String[] splitStudentName(String fullName) {
        if (fullName == null) {
            return new String[]{"", ""};
        }
        String normalized = fullName.trim().replaceAll("\\s+", " ");
        if (normalized.isEmpty()) {
            return new String[]{"", ""};
        }
        int lastSpace = normalized.lastIndexOf(' ');
        if (lastSpace < 0) {
            return new String[]{"", normalized};
        }
        String lastName = normalized.substring(0, lastSpace);
        String firstName = normalized.substring(lastSpace + 1);
        return new String[]{lastName, firstName};
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

    private void applyHeaderBorder(Sheet sheet, CellRangeAddress region) {
        RegionUtil.setBorderTop(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderBottom(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderLeft(BorderStyle.THIN, region, sheet);
        RegionUtil.setBorderRight(BorderStyle.THIN, region, sheet);
    }

    private static class SummaryEntry {
        final String studentId;
        String studentName;
        String className;
        final List<Double> scores;
        final List<String> comments;

        SummaryEntry(String studentId, String studentName, String className, int lecturerCount) {
            this.studentId = studentId;
            this.studentName = studentName;
            this.className = className;
            this.scores = new ArrayList<>(Collections.nCopies(lecturerCount, null));
            this.comments = new ArrayList<>();
        }
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
        final CellStyle cellLeftWrap;
        final CellStyle summaryGpaHeader;
        final CellStyle summaryGpaCell;
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

            cellLeftWrap = workbook.createCellStyle();
            cellLeftWrap.cloneStyleFrom(cellLeft);
            cellLeftWrap.setWrapText(true);

            summaryGpaCell = workbook.createCellStyle();
            summaryGpaCell.cloneStyleFrom(cellCenter);
            summaryGpaCell.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            summaryGpaCell.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            summaryGpaHeader = workbook.createCellStyle();
            summaryGpaHeader.cloneStyleFrom(header);
            summaryGpaHeader.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
            summaryGpaHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);

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

