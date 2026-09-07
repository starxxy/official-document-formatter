import os
import sys
import json
from typing import Dict, List, Optional
from docx import Document

from font_manager import FontManager
from doc_processor import DocProcessor


class OfficialDocFormatter:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                "config",
                "official_doc_config.json"
            )
        
        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)
        
        self.font_manager = FontManager(config_path)
        self.doc_processor = DocProcessor(config_path)
    
    def ensure_fonts_ready(self) -> bool:
        print("\n检查公文字体...")
        fonts_result = self.font_manager.ensure_fonts_available()
        
        all_success = all(info['success'] for info in fonts_result.values())
        
        if all_success:
            print("✓ 所有字体已就绪")
            return True
        else:
            print("\n✗ 部分字体未就绪:")
            for font_name, info in fonts_result.items():
                if not info['success']:
                    print(f"  - {font_name}: {info['message']}")
                    print(self.font_manager.get_manual_install_guide(font_name))
            return False
    
    def get_output_path(self, input_path: str, output_mode: str = None) -> str:
        if output_mode is None:
            output_mode = self.config['output_options']['default_mode']
        
        if output_mode == 'overwrite':
            return input_path
        elif output_mode == 'create_new':
            base, ext = os.path.splitext(input_path)
            return f"{base}_formatted{ext}"
        else:
            return f"{os.path.splitext(input_path)[0]}_formatted.docx"
    
    def format_single_document(self, file_path: str, output_mode: str = None, main_title_index: Optional[int] = None) -> Dict:
        result = {
            'input_file': file_path,
            'output_file': None,
            'success': False,
            'stats': None,
            'validation': None,
            'errors': []
        }
        
        try:
            if not os.path.exists(file_path):
                result['errors'].append(f"文件不存在: {file_path}")
                return result
            
            print(f"\n处理文档: {file_path}")
            
            doc = self.doc_processor.load_document(file_path)
            
            print("应用公文格式...")
            stats = self.doc_processor.process_document(doc, main_title_index=main_title_index)
            result['stats'] = stats
            
            output_path = self.get_output_path(file_path, output_mode)
            result['output_file'] = output_path
            
            print(f"保存文档: {output_path}")
            if self.doc_processor.save_document(doc, output_path):
                result['success'] = True
                print(f"✓ 文档处理完成")
                
                print("\n验证文档格式...")
                validation = self.doc_processor.validate_document_format(doc)
                result['validation'] = validation
                
                if validation['issues']:
                    print("警告: 发现以下问题:")
                    for issue in validation['issues']:
                        print(f"  - {issue['message']}")
            else:
                result['errors'].append("文档保存失败")
                print("✗ 文档保存失败")
                
        except Exception as e:
            error_msg = f"处理文档时出错: {str(e)}"
            result['errors'].append(error_msg)
            print(f"✗ {error_msg}")
        
        return result
    
    def format_batch_documents(self, file_paths: List[str], output_mode: str = None) -> Dict:
        result = {
            'total_files': len(file_paths),
            'successful': 0,
            'failed': 0,
            'results': [],
            'errors': []
        }
        
        error_handling = self.config['error_handling']['batch_mode']
        
        print(f"\n批量处理 {len(file_paths)} 个文档...")
        
        for i, file_path in enumerate(file_paths, 1):
            print(f"\n[{i}/{len(file_paths)}] 处理: {os.path.basename(file_path)}")
            
            single_result = self.format_single_document(file_path, output_mode)
            result['results'].append(single_result)
            
            if single_result['success']:
                result['successful'] += 1
            else:
                result['failed'] += 1
                result['errors'].extend(single_result['errors'])
                
                if error_handling == 'stop_immediately':
                    print("\n遇到错误，停止处理")
                    break
                elif error_handling == 'skip_and_continue':
                    print("跳过此文档，继续处理...")
                    continue
        
        print(f"\n批量处理完成:")
        print(f"  成功: {result['successful']}")
        print(f"  失败: {result['failed']}")
        
        return result
    
    def create_new_document(self, title: str = None, save_path: str = None) -> Dict:
        result = {
            'output_file': save_path,
            'success': False,
            'errors': []
        }
        
        try:
            print("\n创建新公文文档...")
            
            doc = self.doc_processor.create_new_document(title)
            
            if save_path:
                if self.doc_processor.save_document(doc, save_path):
                    result['success'] = True
                    print(f"✓ 文档已创建: {save_path}")
                else:
                    result['errors'].append("文档保存失败")
                    print("✗ 文档保存失败")
            else:
                result['success'] = True
                result['document'] = doc
                print("✓ 文档已创建（内存中）")
                
        except Exception as e:
            error_msg = f"创建文档时出错: {str(e)}"
            result['errors'].append(error_msg)
            print(f"✗ {error_msg}")
        
        return result
    
    def validate_document(self, file_path: str) -> Dict:
        result = {
            'file_path': file_path,
            'validation': None,
            'errors': []
        }
        
        try:
            if not os.path.exists(file_path):
                result['errors'].append(f"文件不存在: {file_path}")
                return result
            
            print(f"\n验证文档: {file_path}")
            
            doc = self.doc_processor.load_document(file_path)
            validation = self.doc_processor.validate_document_format(doc)
            result['validation'] = validation
            
            print("\n验证结果:")
            
            page_valid = all(
                info['valid'] for info in validation['page_format'].values()
            )
            print(f"  页面格式: {'✓ 符合标准' if page_valid else '✗ 不符合标准'}")
            
            if validation['issues']:
                print(f"  发现问题: {len(validation['issues'])} 个")
                for issue in validation['issues']:
                    print(f"    - {issue['message']}")
            else:
                print("  ✓ 未发现问题")
                
        except Exception as e:
            error_msg = f"验证文档时出错: {str(e)}"
            result['errors'].append(error_msg)
            print(f"✗ {error_msg}")
        
        return result


def main():
    print("="*60)
    print("公文格式化工具")
    print("="*60)
    
    formatter = OfficialDocFormatter()
    
    if not formatter.ensure_fonts_ready():
        print("\n警告: 字体未完全就绪，可能影响格式化效果")
        print("建议先安装所需字体后再使用")
    
    print("\n使用示例:")
    print("1. 格式化单个文档:")
    print("   formatter.format_single_document('input.docx', 'create_new')")
    print("\n2. 批量格式化文档:")
    print("   formatter.format_batch_documents(['doc1.docx', 'doc2.docx'])")
    print("\n3. 创建新文档:")
    print("   formatter.create_new_document('公文标题', 'output.docx')")
    print("\n4. 验证文档格式:")
    print("   formatter.validate_document('document.docx')")


if __name__ == "__main__":
    main()
