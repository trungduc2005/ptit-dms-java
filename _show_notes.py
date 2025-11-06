from pathlib import Path
text = Path('src/main/java/com/javaweb/service/EvaluationExportService.java').read_text(encoding='utf-8')
start = text.index('String[] notes')
print(repr(text[start:start+400]))
