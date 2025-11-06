from pathlib import Path
path = Path('src/main/java/com/javaweb/service/EvaluationExportService.java')
text = path.read_text(encoding='utf-8')
start = text.index('String[] notes = {')
end = text.index('};', start)
block = text[start:end+2]
print(block)
