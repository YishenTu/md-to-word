#!/usr/bin/env python3
import argparse
import logging
import os
import stat
import sys
import tempfile
from pathlib import Path

from docx import Document

from src import __version__
from src.core.markdown_preprocessor import MarkdownPreprocessor
from src.core.pandoc_processor import PandocProcessor
from src.core.word_postprocessor import WordPostprocessor
from src.utils.config_validator import ConfigValidator
from src.utils.exceptions import (
    FileProcessingError,
    ImageProcessingError,
    PandocError,
    PathSecurityError,
)


def resolve_output_path(input_path: Path, requested_output: str | None = None) -> Path:
    """Resolve the requested output and guarantee a DOCX file extension."""
    if requested_output:
        output_path = Path(requested_output).expanduser().resolve()
        if output_path.suffix.lower() != '.docx':
            output_path = output_path.with_suffix('.docx')
        return output_path
    return input_path.with_suffix('.docx')


def convert_document(input_path: Path, output_path: Path) -> Path:
    """Convert one Markdown file and atomically publish a validated DOCX."""
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_mode = stat.S_IMODE(output_path.stat().st_mode) if output_path.exists() else 0o644

    preprocessor = MarkdownPreprocessor()
    preprocessed_data = preprocessor.preprocess_file(str(input_path))
    pandoc_processor = PandocProcessor()
    if not pandoc_processor.check_pandoc_available():
        raise PandocError('Pandoc is required but was not found')

    staging_file = tempfile.NamedTemporaryFile(
        prefix=f'{output_path.stem}-',
        suffix='.docx',
        dir=output_path.parent,
        delete=False,
    )
    staging_path = Path(staging_file.name)
    staging_file.close()

    try:
        pandoc_processor.convert_markdown_to_docx(
            preprocessed_data['content'],
            str(staging_path),
            title=None,
            extra_args=['--resource-path', str(input_path.parent)],
        )

        postprocessor = WordPostprocessor()
        postprocessor.apply_formatting(
            str(staging_path),
            preprocessed_data,
            source_dir=input_path.parent,
        )

        # Reopen the package before publication so corrupt or partially written
        # files never replace a previously valid document.
        try:
            Document(str(staging_path))
        except Exception as error:
            raise FileProcessingError(f'Generated DOCX failed validation: {error}') from error

        staging_path.chmod(output_mode)
        os.replace(staging_path, output_path)
        return output_path
    finally:
        try:
            staging_path.unlink(missing_ok=True)
        except OSError as error:
            logging.warning('Failed to remove staging file %s: %s', staging_path, error)


def main():
    """主程序入口"""
    parser = argparse.ArgumentParser(
        description='将Markdown文件转换为符合GB/T 9704-2012标准的Word公文格式',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用示例:
  %(prog)s input.md                    # 输出到 input.docx
  %(prog)s input.md -o output.docx     # 指定输出文件
  %(prog)s input.md --output output.docx
        """,
    )

    parser.add_argument('input', nargs='?', help='输入的Markdown文件路径')

    parser.add_argument('-o', '--output', help='输出的Word文件路径（默认为输入文件名.docx）')

    parser.add_argument('--version', action='version', version=f'%(prog)s {__version__}')

    parser.add_argument('--check-config', action='store_true', help='仅检查配置而不进行转换')

    parser.add_argument('--skip-validation', action='store_true', help='跳过配置验证（不推荐）')

    parser.add_argument('-v', '--verbose', action='store_true', help='启用详细日志输出（用于调试）')

    parser.add_argument('--force', action='store_true', help='非交互模式下允许覆盖已存在的输出文件，不进行询问')

    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[logging.StreamHandler(sys.stderr)],
    )

    # 如果只是检查配置
    if args.check_config:
        print('检查配置...\n')
        validator = ConfigValidator()
        is_valid, results = validator.validate_all()
        validator.print_results(results)
        sys.exit(0 if is_valid else 1)

    # 运行配置验证（除非跳过）
    if not args.skip_validation:
        validator = ConfigValidator()
        is_valid, results = validator.validate_all()

        # 只在有错误时显示验证结果
        if not is_valid:
            print('配置验证失败：')
            validator.print_results(results)
            print('\n请修复以上错误后再运行转换，或使用 --skip-validation 跳过验证（不推荐）')
            sys.exit(1)

    # 检查是否提供了输入文件
    if not args.input:
        parser.error('必须提供输入文件路径')

    # 输入验证
    try:
        # 验证路径安全性
        input_path = Path(args.input).expanduser().resolve()

        # 检查文件是否存在
        if not input_path.exists():
            print(f'错误：文件不存在: {args.input}', file=sys.stderr)
            sys.exit(1)

        # 检查是否为文件（不是目录）
        if not input_path.is_file():
            print(f'错误：不是文件: {args.input}', file=sys.stderr)
            sys.exit(1)

        # 检查文件扩展名
        valid_extensions = ['.md', '.markdown', '.mdown', '.mkd', '.mdwn']
        if input_path.suffix.lower() not in valid_extensions:
            print(f'错误：需要Markdown文件 ({", ".join(valid_extensions)})', file=sys.stderr)
            sys.exit(1)

        # 检查文件是否可读
        if not os.access(input_path, os.R_OK):
            print(f'错误：无法读取文件: {args.input}', file=sys.stderr)
            sys.exit(1)

        # 检查文件大小（限制为100MB）
        max_size_mb = 100
        file_size_mb = input_path.stat().st_size / (1024 * 1024)
        if file_size_mb > max_size_mb:
            print(f'错误：文件太大 ({file_size_mb:.1f}MB > {max_size_mb}MB)', file=sys.stderr)
            sys.exit(1)

    except PathSecurityError as e:
        print(f'安全错误：{e}', file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f'输入验证错误：{e}', file=sys.stderr)
        sys.exit(1)

    # 验证和设置输出路径
    try:
        output_path = resolve_output_path(input_path, args.output)

        # 检查输出目录是否可写
        output_dir = output_path.parent
        if output_dir.exists() and not os.access(output_dir, os.W_OK):
            # 使用相对路径显示
            try:
                rel_output_dir = os.path.relpath(output_dir)
            except ValueError:
                rel_output_dir = str(output_dir)
            print(f'错误：无法写入目录: {rel_output_dir}', file=sys.stderr)
            sys.exit(1)

        # 确保输出目录存在
        output_dir.mkdir(parents=True, exist_ok=True)

        # 如果输出文件已存在，依据 --force 决定是否覆盖
        if output_path.exists():
            if not args.force:
                # 使用相对路径显示
                try:
                    rel_output_path = os.path.relpath(output_path)
                except ValueError:
                    rel_output_path = str(output_path)
                response = input(f'文件已存在: {rel_output_path}，覆盖？(y/N): ')
                if response.lower() != 'y':
                    print('已取消操作')
                    sys.exit(0)

    except PathSecurityError as e:
        print(f'安全错误：{e}', file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f'输出路径错误：{e}', file=sys.stderr)
        sys.exit(1)

    try:
        convert_document(input_path, output_path)

        # 验证输出文件是否成功创建
        if not output_path.exists():
            raise FileProcessingError('输出文件未成功创建')

        # 显示文件大小信息
        output_size_mb = output_path.stat().st_size / (1024 * 1024)
        # 使用相对路径显示
        try:
            rel_output_path = os.path.relpath(output_path)
        except ValueError:
            rel_output_path = str(output_path)
        print(f'完成: {rel_output_path} ({output_size_mb:.2f}MB)')

    except FileProcessingError as e:
        print(f'错误: {e}', file=sys.stderr)
        logging.error(f'FileProcessingError: {e}', exc_info=True)
        sys.exit(1)
    except PandocError as e:
        print(f'Pandoc错误: {e}', file=sys.stderr)
        logging.error(f'PandocError: {e}', exc_info=True)
        sys.exit(1)
    except ImageProcessingError as e:
        print(f'Image error: {e}', file=sys.stderr)
        logging.error('ImageProcessingError: %s', e, exc_info=True)
        sys.exit(1)
    except PathSecurityError as e:
        print(f'路径错误: {e}', file=sys.stderr)
        logging.error(f'PathSecurityError: {e}', exc_info=True)
        sys.exit(1)
    except KeyboardInterrupt:
        print('\n中断', file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f'未知错误: {e}', file=sys.stderr)
        logging.error(f'Unexpected error: {e}', exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()
