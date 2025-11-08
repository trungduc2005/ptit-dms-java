package com.javaweb.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.*;
import java.util.stream.Collectors;

/**
 * DTO cho payload export phiếu chấm của hội đồng.
 * Dùng CouncilEvaluationDto.Root làm kiểu @RequestBody.
 */
public class CouncilEvaluationDto {

    /* -------- ROOT -------- */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Root {
        private EvaluationForm evaluationForm;
        private List<Lecturer> lecturers;

        public Root() {}
        public Root(EvaluationForm evaluationForm, List<Lecturer> lecturers) {
            this.evaluationForm = evaluationForm;
            this.lecturers = lecturers;
        }
        public EvaluationForm getEvaluationForm() { return evaluationForm; }
        public void setEvaluationForm(EvaluationForm evaluationForm) { this.evaluationForm = evaluationForm; }
        public List<Lecturer> getLecturers() { return lecturers; }
        public void setLecturers(List<Lecturer> lecturers) { this.lecturers = lecturers; }
    }

    /* -------- EVALUATION FORM (siÃªu dá»¯ liá»‡u cá»§a phiáº¿u) -------- */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class EvaluationForm {
        private String title;
        private String evaluationId;
        private String academicYear;
        private String formKey;
        private String evaluatorRole;
        private String description;
        private List<Indicator> indicators;

        public EvaluationForm() {}
        public EvaluationForm(String title, String evaluationId, String academicYear,
                              String formKey, String evaluatorRole, String description,
                              List<Indicator> indicators) {
            this.title = title;
            this.evaluationId = evaluationId;
            this.academicYear = academicYear;
            this.formKey = formKey;
            this.evaluatorRole = evaluatorRole;
            this.description = description;
            this.indicators = indicators;
        }
        public String getTitle() { return title; }
        public void setTitle(String title) { this.title = title; }
        public String getEvaluationId() { return evaluationId; }
        public void setEvaluationId(String evaluationId) { this.evaluationId = evaluationId; }
        public String getAcademicYear() { return academicYear; }
        public void setAcademicYear(String academicYear) { this.academicYear = academicYear; }
        public String getFormKey() { return formKey; }
        public void setFormKey(String formKey) { this.formKey = formKey; }
        public String getEvaluatorRole() { return evaluatorRole; }
        public void setEvaluatorRole(String evaluatorRole) { this.evaluatorRole = evaluatorRole; }
        public String getDescription() { return description; }
        public void setDescription(String description) { this.description = description; }
        public List<Indicator> getIndicators() { return indicators; }
        public void setIndicators(List<Indicator> indicators) { this.indicators = indicators; }
    }

    /* -------- CLO & PI -------- */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Indicator {
        private String cloId;
        private String cloName;
        private String cloDescription;
        private List<Pi> pis;
        private Double weight; // trá»ng sá»‘ cá»§a CLO

        public Indicator() {}
        public Indicator(String cloId, String cloName, String cloDescription,
                         List<Pi> pis, Double weight) {
            this.cloId = cloId;
            this.cloName = cloName;
            this.cloDescription = cloDescription;
            this.pis = pis;
            this.weight = weight;
        }
        public String getCloId() { return cloId; }
        public void setCloId(String cloId) { this.cloId = cloId; }
        public String getCloName() { return cloName; }
        public void setCloName(String cloName) { this.cloName = cloName; }
        public String getCloDescription() { return cloDescription; }
        public void setCloDescription(String cloDescription) { this.cloDescription = cloDescription; }
        public List<Pi> getPis() { return pis; }
        public void setPis(List<Pi> pis) { this.pis = pis; }
        public Double getWeight() { return weight; }
        public void setWeight(Double weight) { this.weight = weight; }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Pi {
        private String cloPisId;
        private String cloPisName;
        private String cloPisDescription;
        private Double cloPisWeight; // trá»ng sá»‘ PI trong CLO

        public Pi() {}
        public Pi(String cloPisId, String cloPisName, String cloPisDescription, Double cloPisWeight) {
            this.cloPisId = cloPisId;
            this.cloPisName = cloPisName;
            this.cloPisDescription = cloPisDescription;
            this.cloPisWeight = cloPisWeight;
        }
        public String getCloPisId() { return cloPisId; }
        public void setCloPisId(String cloPisId) { this.cloPisId = cloPisId; }
        public String getCloPisName() { return cloPisName; }
        public void setCloPisName(String cloPisName) { this.cloPisName = cloPisName; }
        public String getCloPisDescription() { return cloPisDescription; }
        public void setCloPisDescription(String cloPisDescription) { this.cloPisDescription = cloPisDescription; }
        public Double getCloPisWeight() { return cloPisWeight; }
        public void setCloPisWeight(Double cloPisWeight) { this.cloPisWeight = cloPisWeight; }
    }

    /* -------- LECTURER & EVALUATIONS -------- */
    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Lecturer {
        private String lecturerId;
        private String lecturerName;
        private String role;
        private String department;
        private List<StudentEvaluation> evaluations;

        public Lecturer() {}
        public Lecturer(String lecturerId, String lecturerName, String role,
                        List<StudentEvaluation> evaluations) {
            this(lecturerId, lecturerName, role, null, evaluations);
        }
        public Lecturer(String lecturerId, String lecturerName, String role,
                        String department, List<StudentEvaluation> evaluations) {
            this.lecturerId = lecturerId;
            this.lecturerName = lecturerName;
            this.role = role;
            this.department = department;
            this.evaluations = evaluations;
        }
        public String getLecturerId() { return lecturerId; }
        public void setLecturerId(String lecturerId) { this.lecturerId = lecturerId; }
        public String getLecturerName() { return lecturerName; }
        public void setLecturerName(String lecturerName) { this.lecturerName = lecturerName; }
        public String getRole() { return role; }
        public void setRole(String role) { this.role = role; }
        public String getDepartment() { return department; }
        public void setDepartment(String department) { this.department = department; }
        public List<StudentEvaluation> getEvaluations() { return evaluations; }
        public void setEvaluations(List<StudentEvaluation> evaluations) { this.evaluations = evaluations; }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class StudentEvaluation {
        private String studentId;
        private String studentName;
        private String className;
        private String comment;
        private Evaluation evaluations; // object con

        public StudentEvaluation() {}
        public StudentEvaluation(String studentId, String className, Evaluation evaluations) {
            this(studentId, null, className, evaluations, null);
        }
        public StudentEvaluation(String studentId, String studentName, String className, Evaluation evaluations) {
            this(studentId, studentName, className, evaluations, null);
        }
        public StudentEvaluation(String studentId, String studentName, String className,
                                 Evaluation evaluations, String comment) {
            this.studentId = studentId;
            this.studentName = studentName;
            this.className = className;
            this.comment = comment;
            this.evaluations = evaluations;
        }
        public String getStudentId() { return studentId; }
        public void setStudentId(String studentId) { this.studentId = studentId; }
        public String getStudentName() { return studentName; }
        public void setStudentName(String studentName) { this.studentName = studentName; }
        public String getClassName() { return className; }
        public void setClassName(String className) { this.className = className; }
        public String getComment() { return comment; }
        public void setComment(String comment) { this.comment = comment; }
        public Evaluation getEvaluations() { return evaluations; }
        public void setEvaluations(Evaluation evaluations) { this.evaluations = evaluations; }

        /**
 * DTO cho payload export phiếu chấm của hội đồng.
 * Dùng CouncilEvaluationDto.Root làm kiểu @RequestBody.
 */
        public Map<String, Double> scoreMap() {
            if (evaluations == null || evaluations.getScores() == null) return Collections.emptyMap();
            return evaluations.getScores().stream()
                    .filter(s -> s.getPiId() != null)
                    .collect(Collectors.toMap(Score::getPiId, Score::getScore, (a,b)->a));
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Evaluation {
        private String evaluationId;
        private String evaluationTitle;
        private List<Score> scores;
        private Double totalScore;

        public Evaluation() {}
        public Evaluation(String evaluationId, String evaluationTitle,
                          List<Score> scores, Double totalScore) {
            this.evaluationId = evaluationId;
            this.evaluationTitle = evaluationTitle;
            this.scores = scores;
            this.totalScore = totalScore;
        }
        public String getEvaluationId() { return evaluationId; }
        public void setEvaluationId(String evaluationId) { this.evaluationId = evaluationId; }
        public String getEvaluationTitle() { return evaluationTitle; }
        public void setEvaluationTitle(String evaluationTitle) { this.evaluationTitle = evaluationTitle; }
        public List<Score> getScores() { return scores; }
        public void setScores(List<Score> scores) { this.scores = scores; }
        public Double getTotalScore() { return totalScore; }
        public void setTotalScore(Double totalScore) { this.totalScore = totalScore; }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Score {
        private String piId;   // vÃ­ dá»¥: "C3.2."
        private Double score;  // vÃ­ dá»¥: 9.5

        public Score() {}
        public Score(String piId, Double score) {
            this.piId = piId;
            this.score = score;
        }
        public String getPiId() { return piId; }
        public void setPiId(String piId) { this.piId = piId; }
        public Double getScore() { return score; }
        public void setScore(Double score) { this.score = score; }
    }
}




