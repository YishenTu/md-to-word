import logging
import os
import subprocess
import tempfile

from docx import Document
from docx.document import Document as DocxDocument

from ..config import DocumentConfig
from ..utils.exceptions import PandocError
from ..utils.path_validator import validate_safe_path


class PandocProcessor:
    """使用Pandoc进行Markdown到Word的转换处理器"""

    def __init__(self):
        self.temp_files: list[str] = []  # 用于跟踪临时文件以便清理
        self.image_config = DocumentConfig.IMAGE_CONFIG
        self.pandoc_config = DocumentConfig.PANDOC_CONFIG

    def convert_markdown_to_docx(
        self,
        markdown_content: str,
        output_path: str,
        title: str | None = None,
        extra_args: list[str] | None = None,
    ) -> str:
        """
        将预处理后的Markdown内容转换为Word文档

        Args:
            markdown_content: 预处理后的Markdown内容
            output_path: 输出的Word文档路径
            title: 文档标题（可选）
            extra_args: 额外的pandoc参数

        Returns:
            生成的Word文档路径
        """
        # 验证输出路径安全性
        safe_output_path = validate_safe_path(output_path)

        try:
            # 创建临时Markdown文件
            with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', suffix='.md', delete=False) as temp_md:
                # 如果提供了标题，添加到内容前面
                if title:
                    # 使用一级标题格式添加文档标题
                    content_with_title = f'# {title}\n\n{markdown_content}'
                else:
                    content_with_title = markdown_content

                temp_md.write(content_with_title)
                temp_md_path = temp_md.name
                self.temp_files.append(temp_md_path)

            # 设置pandoc转换参数，启用TeX数学公式支持
            pandoc_args = self._get_pandoc_args()
            pandoc_args.extend(
                [
                    '-f',
                    'markdown-smart+tex_math_dollars+tex_math_single_backslash-implicit_figures',
                ]
            )
            if extra_args:
                pandoc_args.extend(extra_args)

            cmd = ['pandoc', temp_md_path, '-t', 'docx', '-o', str(safe_output_path)]
            for arg in pandoc_args:
                cmd.append(arg)

            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode != 0:
                raise PandocError(f'Pandoc命令执行失败: {result.stderr}')

            return str(safe_output_path)

        except PandocError:
            raise
        except OSError as error:
            raise PandocError(f'文件操作失败: {error}') from error
        except Exception as error:
            raise PandocError(f'Pandoc处理时发生未知错误: {error}') from error
        finally:
            # 清理临时文件
            self._cleanup_temp_files()

    def _get_pandoc_args(self) -> list[str]:
        """获取pandoc转换参数"""
        args = [
            # 数学公式支持
            f'--{self.pandoc_config["math_method"]}',
        ]

        # 添加配置文件中的额外参数
        if 'extra_args' in self.pandoc_config:
            args.extend(self.pandoc_config['extra_args'])

        return args

    def _cleanup_temp_files(self):
        """清理临时文件"""
        for temp_file in self.temp_files:
            try:
                if os.path.exists(temp_file):
                    os.unlink(temp_file)
            except OSError as e:
                # 记录清理错误但不中断程序
                logging.debug(f'清理临时文件失败 {temp_file}: {e}')
        self.temp_files.clear()

    def load_docx_for_postprocessing(self, docx_path: str) -> DocxDocument:
        """
        加载pandoc生成的docx文件以便进行后处理

        Args:
            docx_path: Word文档路径

        Returns:
            python-docx Document对象
        """
        # 验证路径安全性
        safe_path = validate_safe_path(docx_path)

        try:
            return Document(str(safe_path))
        except Exception as error:
            raise PandocError(f'无法加载Word文档: {error}') from error

    def check_pandoc_available(self) -> bool:
        """检查pandoc是否可用"""
        try:
            result = subprocess.run(
                ['pandoc', '--version'],
                capture_output=True,
                text=True,
            )
            return result.returncode == 0
        except (OSError, FileNotFoundError):
            return False
