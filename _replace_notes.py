from pathlib import Path
path = Path('src/main/java/com/javaweb/service/EvaluationExportService.java')
text = path.read_text(encoding='utf-8')
old_fragment = '''String[] notes = {
                "Ghi ch\xfa:",
                "CLO 3 Thi\u1ebft k\u1ebf ph\u1ea7n c\u1ee9ng v\xe0 ph\u1ea7n m\u1ec1m, ph\xe2n t\xedch d\u1eef li\u1ec7u \u0111\u1ec3 \u0111\xe1nh gi\xe1 hi\u1ec7u qu\u1ea3 ho\u1ea1t \u0111\u1ed9ng c\u1ee7a h\u1ec7 th\u1ed1ng \u0111i\u1ec7n t\u1eed \u0111i\u1ec7n t\u1eed.",
                "C3.3 Ti\u1ebfn h\xe0nh \u0111\u01b0\u1ee3c c\xe1c th\xed nghi\u1ec7m, c\u0169ng nh\u01b0 ph\xe2n t\xedch, \u0111\xe1nh gi\xe1 v\xe0 di\u1ec5n gi\u1ea3i c\xe1c k\u1ebft qu\u1ea3 th\xed nghi\u1ec7m.",
                "C4.0 Th\u1ec3 hi\u1ec7n \u0111\u01b0\u1ee3c \u0111\u1ea1o \u0111\u1ee9c v\xe0 tr\xe1ch nhi\u1ec7m ngh\u1ec1 ngh\u1ec7p trong qu\xe1 tr\xecnh tri\u1ec3n khai c\xe1c h\u1ec7 th\u1ed1ng \u0111i\u1ec7n.",
                "C4.2 Gi\u1ea3i th\xedch \u0111\u01b0\u1ee3c t\xe1c \u0111\u1ed9ng c\u1ee7a k\u1ebft qu\u1ea3 nghi\u00ean c\u1ee9u \u0111\u1ed1i v\u1edbi c\u1ed9ng \u0111\u1ed3ng, x\xe3 h\u1ed9i, ho\u1eb7c ng\xe0nh ngh\u1ec1.",
                "C5.3 Hi\u1ec7u qu\u1ea3 gi\u1ea3i quy\u1ebft v\u1ea5n \u0111\u1ec1 c\u1ee7a nh\xf3m.",
                "CLO 6 V\u1eadn d\u1ee5ng k\u1ef9 n\u0103ng giao ti\u1ebfp trong ng\xe0nh \u0111i\u1ec7n - \u0111i\u1ec7n t\u1eed.",
                "C6.3 Kh\u1ea3 n\u0103ng thuy\u1ebft tr\xednh.",
                "C6.4 Kh\u1ea3 n\u0103ng giao ti\u1ebfp \u0111\u1ed1i tho\u1ea1i v\xe0 tr\u1ea3 l\u1eddi c\u00e1c c\u00e2u h\u1ecfi c\u1ee7a h\u1ed9i \u0111\u1ed3ng."
        };'''
new_fragment = '''String[] notes = {
                "Ghi chú:",
                "CLO 3 Thiết kế phần cứng và phần mềm, phân tích dữ liệu để đánh giá hiệu quả hoạt động của hệ thống điện tử điện tử.",
                "C3.3 Tiến hành được các thí nghiệm, cũng như phân tích, đánh giá và diễn giải các kết quả thí nghiệm.",
                "C4.0 Thể hiện được đạo đức và trách nhiệm nghề nghiệp trong quá trình triển khai các hệ thống điện.",
                "C4.2 Giải thích được tác động của kết quả nghiên cứu đối với cộng đồng, xã hội, hoặc ngành nghề.",
                "C5.3 Hiệu quả giải quyết vấn đề của nhóm.",
                "CLO 6 Vận dụng kỹ năng giao tiếp trong ngành điện - điện tử.",
                "C6.3 Khả năng thuyết trình.",
                "C6.4 Khả năng giao tiếp đối thoại và trả lời các câu hỏi của hội đồng."
        };'''
if old_fragment not in text:
    raise SystemExit('old fragment not found')
text = text.replace(old_fragment, new_fragment)
old_title = 'setCell(signerTitle, signatureStartColumn, "NG\\u01af\\u1edcI \\u0110\\u00c1NH GI\\u00c1", styles.boldLeft);'
new_title = 'setCell(signerTitle, signatureStartColumn, "NGƯỜI ĐÁNH GIÁ", styles.boldLeft);'
text = text.replace(old_title, new_title)
path.write_text(text, encoding='utf-8')
