package com.javaweb.dto;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * DTO cho luồng xuất file đánh giá của giảng viên hướng dẫn.
 * Dùng GuiderEvaluationDto.Root làm kiểu @RequestBody.
 */
public class GuiderEvaluationDto {

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Root {
        private Boolean success;
        private List<EvaluationForm> evaluationForm;
        private List<Student> students;

        public Boolean getSuccess() {
            return success;
        }

        public void setSuccess(Boolean success) {
            this.success = success;
        }

        public List<EvaluationForm> getEvaluationForm() {
            return evaluationForm;
        }

        public void setEvaluationForm(List<EvaluationForm> evaluationForm) {
            this.evaluationForm = evaluationForm;
        }

        public List<Student> getStudents() {
            return students;
        }

        public void setStudents(List<Student> students) {
            this.students = students;
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class EvaluationForm {
        private String title;
        private String evaluationId;
        private String academicYear;
        private String formKey;
        private String reportWeek;
        private String evaluatorRole;
        private String description;
        private List<Indicator> indicators;

        public String getTitle() {
            return title;
        }

        public void setTitle(String title) {
            this.title = title;
        }

        public String getEvaluationId() {
            return evaluationId;
        }

        public void setEvaluationId(String evaluationId) {
            this.evaluationId = evaluationId;
        }

        public String getAcademicYear() {
            return academicYear;
        }

        public void setAcademicYear(String academicYear) {
            this.academicYear = academicYear;
        }

        public String getFormKey() {
            return formKey;
        }

        public void setFormKey(String formKey) {
            this.formKey = formKey;
        }

        public String getReportWeek() {
            return reportWeek;
        }

        public void setReportWeek(String reportWeek) {
            this.reportWeek = reportWeek;
        }

        public String getEvaluatorRole() {
            return evaluatorRole;
        }

        public void setEvaluatorRole(String evaluatorRole) {
            this.evaluatorRole = evaluatorRole;
        }

        public String getDescription() {
            return description;
        }

        public void setDescription(String description) {
            this.description = description;
        }

        public List<Indicator> getIndicators() {
            return indicators;
        }

        public void setIndicators(List<Indicator> indicators) {
            this.indicators = indicators;
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Indicator {
        private String cloId;
        private String cloName;
        private String cloDescription;
        private List<Pi> pis;
        private Double weight;

        public String getCloId() {
            return cloId;
        }

        public void setCloId(String cloId) {
            this.cloId = cloId;
        }

        public String getCloName() {
            return cloName;
        }

        public void setCloName(String cloName) {
            this.cloName = cloName;
        }

        public String getCloDescription() {
            return cloDescription;
        }

        public void setCloDescription(String cloDescription) {
            this.cloDescription = cloDescription;
        }

        public List<Pi> getPis() {
            return pis;
        }

        public void setPis(List<Pi> pis) {
            this.pis = pis;
        }

        public Double getWeight() {
            return weight;
        }

        public void setWeight(Double weight) {
            this.weight = weight;
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Pi {
        private String cloPisId;
        private String cloPisName;
        private String cloPisDescription;
        private Double cloPisWeight;

        public String getCloPisId() {
            return cloPisId;
        }

        public void setCloPisId(String cloPisId) {
            this.cloPisId = cloPisId;
        }

        public String getCloPisName() {
            return cloPisName;
        }

        public void setCloPisName(String cloPisName) {
            this.cloPisName = cloPisName;
        }

        public String getCloPisDescription() {
            return cloPisDescription;
        }

        public void setCloPisDescription(String cloPisDescription) {
            this.cloPisDescription = cloPisDescription;
        }

        public Double getCloPisWeight() {
            return cloPisWeight;
        }

        public void setCloPisWeight(Double cloPisWeight) {
            this.cloPisWeight = cloPisWeight;
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Student {
        private String studentId;
        private String studentName;
        private String avatarUrl;
        private String studentClassName;
        private String role;
        private String guiderName;
        private String projectName;
        private List<StudentEvaluation> evaluations;

        public String getStudentId() {
            return studentId;
        }

        public void setStudentId(String studentId) {
            this.studentId = studentId;
        }

        public String getStudentName() {
            return studentName;
        }

        public void setStudentName(String studentName) {
            this.studentName = studentName;
        }

        public String getAvatarUrl() {
            return avatarUrl;
        }

        public void setAvatarUrl(String avatarUrl) {
            this.avatarUrl = avatarUrl;
        }

        public String getStudentClassName() {
            return studentClassName;
        }

        public void setStudentClassName(String studentClassName) {
            this.studentClassName = studentClassName;
        }

        public String getRole() {
            return role;
        }

        public void setRole(String role) {
            this.role = role;
        }

        public String getGuiderName() {
            return guiderName;
        }

        public void setGuiderName(String guiderName) {
            this.guiderName = guiderName;
        }

        public String getProjectName() {
            return projectName;
        }

        public void setProjectName(String projectName) {
            this.projectName = projectName;
        }

        public List<StudentEvaluation> getEvaluations() {
            return evaluations;
        }

        public void setEvaluations(List<StudentEvaluation> evaluations) {
            this.evaluations = evaluations;
        }

        public Map<String, StudentEvaluation> evaluationMap() {
            if (evaluations == null) {
                return Collections.emptyMap();
            }
            return evaluations.stream()
                    .filter(e -> e.getEvaluationId() != null)
                    .collect(Collectors.toMap(StudentEvaluation::getEvaluationId, e -> e, (a, b) -> a));
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class StudentEvaluation {
        private String evaluationId;
        private String evaluationTitle;
        private List<Score> scores;

        public String getEvaluationId() {
            return evaluationId;
        }

        public void setEvaluationId(String evaluationId) {
            this.evaluationId = evaluationId;
        }

        public String getEvaluationTitle() {
            return evaluationTitle;
        }

        public void setEvaluationTitle(String evaluationTitle) {
            this.evaluationTitle = evaluationTitle;
        }

        public List<Score> getScores() {
            return scores;
        }

        public void setScores(List<Score> scores) {
            this.scores = scores;
        }

        public Map<String, Double> scoreMap() {
            if (scores == null) {
                return Collections.emptyMap();
            }
            return scores.stream()
                    .filter(s -> s.getPiId() != null && s.getScore() != null)
                    .collect(Collectors.toMap(Score::getPiId, Score::getScore, (a, b) -> a));
        }
    }

    @JsonIgnoreProperties(ignoreUnknown = true)
    public static class Score {
        private String piId;
        private Double score;

        public String getPiId() {
            return piId;
        }

        public void setPiId(String piId) {
            this.piId = piId;
        }

        public Double getScore() {
            return score;
        }

        public void setScore(Double score) {
            this.score = score;
        }
    }
}
