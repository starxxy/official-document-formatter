import re
import json
import os
from typing import Dict, List, Optional, Tuple
from docx import Document
from docx.text.paragraph import Paragraph
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT


class TitleDetector:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                "config",
                "official_doc_config.json"
            )
        
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.title_patterns = self.config['title_patterns']
        self.compiled_patterns = {}
        for level, pattern_info in self.title_patterns.items():
            pattern = pattern_info['pattern']
            self.compiled_patterns[level] = re.compile(pattern)

        # 标题检测配置（带默认值，兼容旧配置）
        detection = self.config.get('title_detection', {})
        self.large_font_pt = detection.get('large_font_pt', 22)
        self.main_title_max_length = detection.get('main_title_max_length', 60)
        self.first_paragraph_as_title = detection.get('first_paragraph_as_title', True)
        self.centered_title_max_position = detection.get('centered_title_max_position', 5)
    
    def detect_title_level_by_pattern(self, text: str) -> Optional[int]:
        text = text.strip()
        
        for level in ['level1', 'level2', 'level3', 'level4']:
            pattern = self.compiled_patterns[level]
            if pattern.match(text):
                level_num = int(level.replace('level', ''))
                return level_num
        
        return None
    
    def detect_title_level_by_format(self, paragraph: Paragraph) -> Optional[int]:
        if not paragraph.runs:
            return None
        
        run = paragraph.runs[0]
        
        font_size = None
        if run.font.size:
            font_size = run.font.size.pt
        
        is_bold = run.font.bold if run.font.bold is not None else False
        
        alignment = paragraph.paragraph_format.alignment
        
        if font_size and font_size > 16:
            if is_bold:
                return 1
            else:
                return 2
        
        if font_size and font_size > 14 and is_bold:
            return 1
        
        return None
    
    def detect_title_level(self, paragraph: Paragraph) -> int:
        text = paragraph.text.strip()
        
        if not text:
            return 0
        
        pattern_level = self.detect_title_level_by_pattern(text)
        
        format_level = self.detect_title_level_by_format(paragraph)
        
        if pattern_level is not None:
            return pattern_level
        
        if format_level is not None:
            return format_level
        
        return 0
    
    def is_title(self, paragraph: Paragraph) -> bool:
        level = self.detect_title_level(paragraph)
        return level > 0
    
    def is_main_title(self, paragraph: Paragraph, position: int = 0, allow_first_paragraph: bool = True) -> bool:
        """综合判断是否为主标题。

        position: 当前段落在「非空段落」中的序号（从 0 开始），
        用于识别文档首个非空段，并避免把文末居中的落款误判为主标题。
        allow_first_paragraph: 是否启用「首个非空段」启发式；用户手动指定
        主标题索引时应传 False，避免红头/发文机关被误判。
        """
        text = paragraph.text.strip()
        
        if not text:
            return False
        
        if len(text) > self.main_title_max_length:
            return False
        
        font_size = None
        if paragraph.runs:
            run = paragraph.runs[0]
            font_size = run.font.size.pt if run.font.size else None
        
        alignment = paragraph.paragraph_format.alignment
        
        # 1. 大字号（如2号字）→ 主标题
        if font_size and font_size >= self.large_font_pt:
            return True
        
        # 2. 居中对齐且位于文档前部 → 主标题
        #    公文标题常居中标示，限定前部可避免误判文末落款/日期
        if alignment == WD_PARAGRAPH_ALIGNMENT.CENTER and position <= self.centered_title_max_position:
            return True
        
        # 3. 首个非空段落且不带标题编号 → 主标题
        #    兜底无格式标题（如"关于开展工作的通知"直接作为首段）
        if allow_first_paragraph and self.first_paragraph_as_title and position == 0 and self.detect_title_level_by_pattern(text) is None:
            return True
        
        return False
    
    def extract_title_info(self, paragraph: Paragraph, position: int = 0) -> Dict:
        text = paragraph.text.strip()
        level = self.detect_title_level(paragraph)
        is_main = self.is_main_title(paragraph, position)
        
        info = {
            'text': text,
            'level': level,
            'is_title': level > 0,
            'is_main_title': is_main,
            'paragraph_index': None,
            'format_info': {}
        }
        
        if paragraph.runs:
            run = paragraph.runs[0]
            info['format_info'] = {
                'font_size': run.font.size.pt if run.font.size else None,
                'font_name': run.font.name,
                'bold': run.font.bold if run.font.bold is not None else False,
                'alignment': str(paragraph.paragraph_format.alignment)
            }
        
        return info
    
    def scan_document_titles(self, doc: Document) -> List[Dict]:
        titles = []
        
        position = 0
        for i, paragraph in enumerate(doc.paragraphs):
            if not paragraph.text.strip():
                continue
            
            info = self.extract_title_info(paragraph, position)
            info['paragraph_index'] = i
            
            if info['is_title'] or info['is_main_title']:
                titles.append(info)
            
            position += 1
        
        return titles
    
    def get_title_statistics(self, doc: Document) -> Dict:
        titles = self.scan_document_titles(doc)
        
        stats = {
            'total_titles': len(titles),
            'main_title_count': sum(1 for t in titles if t['is_main_title']),
            'level_distribution': {
                'level1': 0,
                'level2': 0,
                'level3': 0,
                'level4': 0
            },
            'titles': titles
        }
        
        for title in titles:
            if not title['is_main_title'] and title['level'] > 0:
                level_key = f"level{title['level']}"
                if level_key in stats['level_distribution']:
                    stats['level_distribution'][level_key] += 1
        
        return stats
    
    def validate_title_hierarchy(self, doc: Document) -> Dict:
        titles = self.scan_document_titles(doc)
        
        issues = []
        prev_level = 0
        
        for i, title in enumerate(titles):
            if title['is_main_title']:
                continue
            
            current_level = title['level']
            
            if current_level > 0:
                if prev_level > 0 and current_level > prev_level + 1:
                    issues.append({
                        'type': 'level_skip',
                        'message': f"标题层级跳跃: 从{prev_level}级跳到{current_level}级",
                        'title': title['text'],
                        'index': title['paragraph_index']
                    })
                
                prev_level = current_level
        
        return {
            'valid': len(issues) == 0,
            'issues': issues,
            'total_titles': len(titles)
        }
    
    def suggest_title_level(self, paragraph: Paragraph, context: Dict = None) -> int:
        text = paragraph.text.strip()
        
        pattern_level = self.detect_title_level_by_pattern(text)
        if pattern_level is not None:
            return pattern_level
        
        format_level = self.detect_title_level_by_format(paragraph)
        if format_level is not None:
            return format_level
        
        if context and 'prev_level' in context:
            return context['prev_level'] + 1
        
        return 0


def main():
    print("公文格式化 - 标题识别测试")
    print("="*60)
    
    detector = TitleDetector()
    
    test_texts = [
        "一、测试标题",
        "（一）二级标题",
        "1. 三级标题",
        "（1）四级标题",
        "这是正文内容",
        "二、另一个一级标题"
    ]
    
    print("\n测试编号格式识别:")
    for text in test_texts:
        level = detector.detect_title_level_by_pattern(text)
        status = f"{level}级标题" if level else "非标题"
        print(f"  '{text}' -> {status}")
    
    print("\n标题识别模块测试完成！")


if __name__ == "__main__":
    main()
