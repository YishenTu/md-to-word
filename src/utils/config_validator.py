#!/usr/bin/env python3
"""
配置验证器 - 检查系统环境和配置是否满足运行要求
"""

import os
import shutil
import subprocess
import sys
from importlib import metadata
from pathlib import Path

from ..config import DocumentConfig


class ConfigValidator:
    """配置验证器类"""

    def __init__(self):
        self.errors: list[str] = []
        self.warnings: list[str] = []
        self.info: list[str] = []

    def validate_all(self) -> tuple[bool, dict[str, list[str]]]:
        """
        运行所有验证检查

        Returns:
            Tuple[bool, Dict]: (是否通过验证, {errors, warnings, info})
        """
        self.errors = []
        self.warnings = []
        self.info = []

        # 运行各项检查
        self._check_python_version()
        self._check_pandoc_installation()
        self._check_python_dependencies()
        self._check_obsidian_paths()
        self._check_image_search_paths()
        self._check_fonts()

        # 返回结果
        is_valid = len(self.errors) == 0
        results = {'errors': self.errors, 'warnings': self.warnings, 'info': self.info}

        return is_valid, results

    def _check_python_version(self):
        """检查Python版本"""
        python_version = sys.version_info
        if python_version.major < 3 or (python_version.major == 3 and python_version.minor < 11):
            self.errors.append(
                f'Python版本过低：当前版本 {python_version.major}.{python_version.minor}，需要 Python 3.11 或更高版本'
            )
        else:
            self.info.append(f'Python版本：{python_version.major}.{python_version.minor}.{python_version.micro}')

    def _check_pandoc_installation(self):
        """检查Pandoc是否已安装"""
        pandoc_path = shutil.which('pandoc')

        if not pandoc_path:
            self.errors.append(
                '未找到Pandoc安装。请安装Pandoc后再使用本工具。\n'
                '  安装说明：https://pandoc.org/installing.html\n'
                '  - macOS: brew install pandoc\n'
                '  - Windows: 下载安装包或使用 choco install pandoc\n'
                '  - Linux: sudo apt-get install pandoc 或 sudo yum install pandoc'
            )
        else:
            # 检查版本
            try:
                result = subprocess.run(['pandoc', '--version'], capture_output=True, text=True, check=True)
                version_line = result.stdout.split('\n')[0]
                self.info.append(f'Pandoc已安装：{version_line}')
                # 使用相对路径显示
                try:
                    rel_pandoc_path = os.path.relpath(pandoc_path)
                except ValueError:
                    rel_pandoc_path = pandoc_path
                self.info.append(f'Pandoc路径：{rel_pandoc_path}')
            except subprocess.CalledProcessError:
                self.warnings.append(f'Pandoc已找到但无法获取版本信息：{pandoc_path}')

    def _check_python_dependencies(self):
        """检查Python依赖包"""
        required_packages = {'docx': ('python-docx', '1.2.0')}

        missing_packages = []

        for module_name, (package_name, required_version) in required_packages.items():
            try:
                __import__(module_name)
                installed_version = metadata.version(package_name)
                requirement = f'{package_name}=={required_version}'
                if installed_version != required_version:
                    self.errors.append(
                        f'依赖包版本不匹配：已安装 {package_name}=={installed_version}，需要 {requirement}'
                    )
                else:
                    self.info.append(f'依赖包已安装：{requirement}')
            except (ImportError, metadata.PackageNotFoundError):
                missing_packages.append(f'{package_name}=={required_version}')

        if missing_packages:
            self.errors.append(
                f'缺少必要的Python依赖包：{", ".join(missing_packages)}\n'
                f'  请运行以下命令安装：\n'
                f'  pip3 install -r requirements.txt\n'
                f'  或者：\n'
                f'  pip3 install {" ".join(missing_packages)}'
            )

    def _check_obsidian_paths(self):
        """检查Obsidian相关路径配置"""
        config = DocumentConfig.OBSIDIAN_CONFIG
        vault_path_value = config['vault_path']
        vault_name = config['vault_name']
        attachments_folder = config['attachments_folder']

        # 检查环境变量配置
        vault_path_env = os.getenv('OBSIDIAN_VAULT_PATH')
        vault_name_env = os.getenv('OBSIDIAN_VAULT_NAME')
        attachments_folder_env = os.getenv('OBSIDIAN_ATTACHMENTS_FOLDER')

        if vault_path_env:
            self.info.append(f'检测到OBSIDIAN_VAULT_PATH环境变量：{vault_path_env}')
            vault_path = Path(vault_path_env)
            if not vault_path.exists():
                self.warnings.append(
                    f'OBSIDIAN_VAULT_PATH指定的路径不存在：{vault_path_env}\n  图片搜索可能无法正常工作'
                )
            else:
                self.info.append(f'Obsidian Vault路径有效：{vault_path_env}')

        if vault_name_env:
            self.info.append(f'检测到OBSIDIAN_VAULT_NAME环境变量：{vault_name_env}')

        if attachments_folder_env:
            self.info.append(f'检测到OBSIDIAN_ATTACHMENTS_FOLDER环境变量：{attachments_folder_env}')

        # 如果配置了完整路径
        if vault_path_value:
            vault_path = Path(vault_path_value)
            if vault_path.exists():
                attachments_path = vault_path / attachments_folder
                if attachments_path.exists():
                    try:
                        rel_attachments_path = os.path.relpath(attachments_path)
                    except ValueError:
                        rel_attachments_path = str(attachments_path)
                    self.info.append(f'Obsidian附件目录存在：{rel_attachments_path}')
                else:
                    try:
                        rel_attachments_path = os.path.relpath(attachments_path)
                    except ValueError:
                        rel_attachments_path = str(attachments_path)
                    self.warnings.append(
                        f'Obsidian附件目录不存在：{rel_attachments_path}\n'
                        f'  如果文档包含Obsidian图片引用，可能无法找到图片'
                    )

        # 自动检测
        elif vault_name:
            # 检测常见路径
            search_locations = [
                (
                    Path.home() / 'Library/Mobile Documents/iCloud~md~obsidian/Documents' / vault_name,
                    'iCloud',
                ),
                (Path.home() / 'Documents' / vault_name, 'Documents'),
                (Path.home() / 'Desktop' / vault_name, 'Desktop'),
            ]

            found = False
            for vault_path, location in search_locations:
                if vault_path.exists():
                    try:
                        rel_vault_path = os.path.relpath(vault_path)
                    except ValueError:
                        rel_vault_path = str(vault_path)
                    self.info.append(f'自动检测到Obsidian Vault（{location}）：{rel_vault_path}')
                    attachments_path = vault_path / attachments_folder
                    if attachments_path.exists():
                        try:
                            rel_attachments_path = os.path.relpath(attachments_path)
                        except ValueError:
                            rel_attachments_path = str(attachments_path)
                        self.info.append(f'Obsidian附件目录存在：{rel_attachments_path}')
                    else:
                        try:
                            rel_attachments_path = os.path.relpath(attachments_path)
                        except ValueError:
                            rel_attachments_path = str(attachments_path)
                        self.warnings.append(
                            f'Obsidian附件目录不存在：{rel_attachments_path}\n'
                            f"  附件文件夹名称可能不是 '{attachments_folder}'"
                        )
                    found = True
                    break

            if not found:
                self.warnings.append(
                    f"未找到Obsidian Vault '{vault_name}'\n"
                    f'  检查了以下位置：iCloud、Documents、Desktop\n'
                    f'  可以通过设置环境变量来指定正确的路径：\n'
                    f'  export OBSIDIAN_VAULT_PATH=/path/to/vault\n'
                    f"  export OBSIDIAN_VAULT_NAME='Your Vault Name'\n"
                    f"  export OBSIDIAN_ATTACHMENTS_FOLDER='Attachments Folder Name'"
                )

    def _check_image_search_paths(self):
        """检查图片搜索路径"""
        search_paths = DocumentConfig.get_image_search_paths()

        existing_paths = []
        missing_paths = []

        for path in search_paths:
            path_obj = Path(path)
            if path_obj.exists():
                existing_paths.append(path)
            else:
                # 只对非相对路径报告缺失
                if path_obj.is_absolute():
                    missing_paths.append(path)

        if existing_paths:
            # 转换为相对路径显示
            rel_paths = []
            for path in existing_paths:
                try:
                    rel_path = os.path.relpath(path)
                except ValueError:
                    rel_path = path
                rel_paths.append(rel_path)
            self.info.append('图片搜索路径（存在的）：\n  ' + '\n  '.join(rel_paths))

        if missing_paths:
            self.info.append('图片搜索路径（不存在的）：\n  ' + '\n  '.join(missing_paths))

        # 检查支持的图片格式
        supported_formats = DocumentConfig.IMAGE_CONFIG.get('supported_formats', [])
        self.info.append(f'支持的图片格式：{", ".join(supported_formats)}')

    def _check_fonts(self):
        """检查字体配置（仅提供信息）"""
        fonts = DocumentConfig.FONTS
        self.info.append('配置的字体：')
        for font_type, font_name in fonts.items():
            self.info.append(f'  {font_type}: {font_name}')

        configured_fonts = '\n'.join(f'  - {font_type}: {font_name}' for font_type, font_name in fonts.items())
        self.warnings.append(
            f'请确保系统中安装了配置的字体：\n{configured_fonts}\n  如果缺少这些字体，Word可能使用默认字体替代'
        )

    def print_results(self, results: dict[str, list[str]]):
        """打印验证结果"""
        # 打印错误
        if results['errors']:
            print('\n错误:')
            for i, error in enumerate(results['errors'], 1):
                print(f'\n{i}. {error}')

        if results['warnings']:
            print('\n警告:')
            for i, warning in enumerate(results['warnings'], 1):
                print(f'\n{i}. {warning}')

        if results['info']:
            print('\n系统信息:')
            for info in results['info']:
                print(f'  - {info}')


def validate_config(print_output: bool = True) -> bool:
    """
    验证配置的便捷函数

    Args:
        print_output: 是否打印输出结果

    Returns:
        bool: 是否通过验证
    """
    validator = ConfigValidator()
    is_valid, results = validator.validate_all()

    if print_output:
        validator.print_results(results)

    return is_valid


def main():
    """独立运行配置验证器"""
    print('检查配置...\n')
    is_valid = validate_config()
    sys.exit(0 if is_valid else 1)


if __name__ == '__main__':
    main()
