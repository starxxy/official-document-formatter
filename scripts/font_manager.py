import os
import sys
import json
import shutil
import ctypes
import platform
from typing import Dict, List, Set, Tuple, Optional

# 注：本模块优先使用 fontTools 读取字体「真实字体族名」做精确检测，
# 避免依赖系统字体文件名（文件名与字体族名往往不一致）。
try:
    from fontTools.ttLib import TTFont
    HAS_FONTTOOLS = True
except ImportError:
    HAS_FONTTOOLS = False

try:
    import winreg
    HAS_WINREG = True
except ImportError:
    HAS_WINREG = False


class FontManager:
    def __init__(self, config_path: str = None):
        if config_path is None:
            config_path = os.path.join(
                os.path.dirname(os.path.dirname(__file__)),
                "config",
                "official_doc_config.json"
            )

        with open(config_path, 'r', encoding='utf-8') as f:
            self.config = json.load(f)

        self.fonts_dir = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "fonts"
        )

        self.system_font_dirs = self._get_system_font_dirs()
        # 缓存系统已安装的字体族名集合（小写），避免重复枚举
        self._installed_families_cache: Optional[Set[str]] = None

    # ------------------------------------------------------------------
    # 系统字体目录
    # ------------------------------------------------------------------
    def _get_system_font_dirs(self) -> List[str]:
        system = platform.system()
        dirs: List[str] = []

        if system == "Windows":
            # 用户级字体目录（无需管理员权限）
            local = os.environ.get("LOCALAPPDATA")
            if local:
                dirs.append(os.path.join(local, "Microsoft", "Windows", "Fonts"))
            # 系统级字体目录（通常需要管理员权限）
            windir = os.environ.get("WINDIR", r"C:\Windows")
            dirs.append(os.path.join(windir, "Fonts"))
        elif system == "Darwin":
            dirs.append(os.path.expanduser("~/Library/Fonts"))
            dirs.append("/Library/Fonts")
            dirs.append("/System/Library/Fonts")
        elif system == "Linux":
            dirs.append(os.path.expanduser("~/.fonts"))
            dirs.append("/usr/share/fonts")
            dirs.append("/usr/local/share/fonts")

        return [d for d in dirs if d and os.path.isdir(d)]

    # ------------------------------------------------------------------
    # 字体族名读取
    # ------------------------------------------------------------------
    @staticmethod
    def _read_font_families(path: str) -> Set[str]:
        """读取单个字体文件的全部 family name（nameID=1），统一转小写。"""
        families: Set[str] = set()
        if not HAS_FONTTOOLS or not os.path.exists(path):
            return families
        try:
            font = TTFont(path, lazy=True, fontNumber=0)
            for record in font["name"].names:
                if record.nameID == 1:
                    try:
                        val = record.toUnicode()
                    except Exception:
                        continue
                    if val and val.strip():
                        families.add(val.strip().lower())
        except Exception:
            # .ttc 集合字体或损坏文件，忽略
            pass
        return families

    def _get_installed_font_families(self) -> Set[str]:
        """枚举系统字体目录，返回已安装字体的族名集合（小写，结果缓存）。"""
        if self._installed_families_cache is not None:
            return self._installed_families_cache

        families: Set[str] = set()
        for d in self.system_font_dirs:
            try:
                entries = os.listdir(d)
            except Exception:
                continue
            for name in entries:
                if name.lower().endswith((".ttf", ".ttc", ".otf")):
                    families |= self._read_font_families(os.path.join(d, name))

        self._installed_families_cache = families
        return families

    # ------------------------------------------------------------------
    # 内置字体定位
    # ------------------------------------------------------------------
    def _required_fonts(self) -> Dict[str, Dict]:
        fonts = self.config.get("required_fonts") or self.config.get("fonts_download") or {}
        return fonts

    def _get_bundled_font_path(self, font_name: str) -> Optional[str]:
        info = self._required_fonts().get(font_name)
        if not info:
            return None
        fn = info.get("bundled_filename") or info.get("filename")
        if not fn:
            return None
        path = os.path.join(self.fonts_dir, fn)
        return path if os.path.exists(path) else None

    def _get_target_filename(self, font_name: str) -> str:
        info = self._required_fonts().get(font_name)
        return info.get("target_filename") or info.get("bundled_filename") or info.get("filename") or (font_name + ".ttf")

    # ------------------------------------------------------------------
    # 检测
    # ------------------------------------------------------------------
    def check_font_installed(self, font_name: str) -> bool:
        """精确检测某个公文字体是否已安装（按真实字体族名匹配）。"""
        if HAS_FONTTOOLS:
            bundled_path = self._get_bundled_font_path(font_name)
            if bundled_path:
                target_families = self._read_font_families(bundled_path)
                installed = self._get_installed_font_families()
                if target_families & installed:
                    return True

        # 降级方案：无 fontTools 时，用注册表显示名匹配
        if self._check_font_in_registry(font_name):
            return True

        # 兜底：按目标文件名精确匹配
        target = self._get_target_filename(font_name).lower()
        for d in self.system_font_dirs:
            try:
                for name in os.listdir(d):
                    if name.lower() == target:
                        return True
            except Exception:
                continue

        return False

    def _check_font_in_registry(self, font_name: str) -> bool:
        """降级检测：读取 Windows 字体注册表（HKLM/HKCU）中的字体显示名。"""
        if not HAS_WINREG or platform.system() != "Windows":
            return False

        keys = [
            (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
            (winreg.HKEY_CURRENT_USER, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"),
        ]
        fams = self._required_fonts().get(font_name, {}).get("family_names", [font_name])
        targets = {f.strip().lower() for f in fams}

        for hive, subkey in keys:
            try:
                with winreg.OpenKey(hive, subkey) as key:
                    i = 0
                    while True:
                        try:
                            display_name = winreg.EnumValue(key, i)[0]
                        except OSError:
                            break
                        # 去掉 "(TrueType)" 等后缀
                        clean = display_name.split("(")[0].strip().lower()
                        if clean in targets:
                            return True
                        i += 1
            except OSError:
                continue
        return False

    def check_all_fonts(self) -> Dict[str, bool]:
        result = {}
        for font_name in self._required_fonts().keys():
            result[font_name] = self.check_font_installed(font_name)
        return result

    # ------------------------------------------------------------------
    # 安装
    # ------------------------------------------------------------------
    def _broadcast_font_change(self) -> None:
        """广播 WM_FONTCHANGE，通知系统字体列表已更新。"""
        if platform.system() != "Windows":
            return
        try:
            WM_FONTCHANGE = 0x001D
            HWND_BROADCAST = 0xFFFF
            ctypes.windll.user32.PostMessageW(HWND_BROADCAST, WM_FONTCHANGE, 0, 0)
        except Exception:
            pass

    def _install_windows(self, font_file_path: str, target_filename: str) -> Tuple[bool, str]:
        if not os.path.exists(font_file_path):
            return False, f"字体文件不存在: {font_file_path}"

        display_name = os.path.splitext(target_filename)[0]
        # 优先使用字体自身的中文族名作为注册表显示名
        families = self._read_font_families(font_file_path)
        # family_names 配置里的中文名优先（当前文件读取顺序可能无序）
        info_value = target_filename

        local_dir = ""
        local = os.environ.get("LOCALAPPDATA")
        if local:
            local_dir = os.path.join(local, "Microsoft", "Windows", "Fonts")

        errors = []

        # 方案一：用户级安装（无需管理员）
        if local_dir:
            try:
                os.makedirs(local_dir, exist_ok=True)
                dest = os.path.join(local_dir, target_filename)
                shutil.copy2(font_file_path, dest)

                if HAS_WINREG:
                    self._write_registry(
                        winreg.HKEY_CURRENT_USER,
                        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
                        display_name,
                        target_filename,
                    )
                self._broadcast_font_change()
                return True, f"字体已安装（用户级）: {target_filename}"
            except Exception as e:
                errors.append(f"用户级安装失败: {e}")

        # 方案二：系统级安装（需管理员）
        windir = os.environ.get("WINDIR", r"C:\Windows")
        sys_dir = os.path.join(windir, "Fonts")
        try:
            dest = os.path.join(sys_dir, target_filename)
            shutil.copy2(font_file_path, dest)

            if HAS_WINREG:
                self._write_registry(
                    winreg.HKEY_LOCAL_MACHINE,
                    r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts",
                    display_name,
                    target_filename,
                )
            self._broadcast_font_change()
            return True, f"字体已安装（系统级）: {target_filename}"
        except Exception as e:
            errors.append(f"系统级安装失败: {e}")

        return False, "; ".join(errors)

    def _write_registry(self, hive, subkey: str, display_name: str, value: str) -> None:
        try:
            with winreg.CreateKey(hive, subkey) as key:
                reg_name = display_name + " (TrueType)"
                winreg.SetValueEx(key, reg_name, 0, winreg.REG_SZ, value)
        except Exception as e:
            # 注册表写入失败不阻断（文件已复制，部分场景仍可用）
            print(f"  [警告] 注册表写入失败: {e}")

    def _install_linux(self, font_file_path: str, target_filename: str) -> Tuple[bool, str]:
        # 优先装到用户目录，无需管理员
        user_dir = os.path.expanduser("~/.fonts")
        try:
            os.makedirs(user_dir, exist_ok=True)
            shutil.copy2(font_file_path, os.path.join(user_dir, target_filename))
        except Exception as e:
            return False, f"字体安装失败: {e}"

        try:
            subprocess = __import__("subprocess")
            subprocess.run(["fc-cache", "-f"], check=False)
        except Exception:
            pass
        return True, f"字体已安装: {target_filename}"

    def _install_mac(self, font_file_path: str, target_filename: str) -> Tuple[bool, str]:
        user_dir = os.path.expanduser("~/Library/Fonts")
        try:
            os.makedirs(user_dir, exist_ok=True)
            shutil.copy2(font_file_path, os.path.join(user_dir, target_filename))
            return True, f"字体已安装: {target_filename}"
        except Exception as e:
            return False, f"字体安装失败: {e}"

    def install_font_file(self, font_file_path: str, target_filename: str) -> Tuple[bool, str]:
        system = platform.system()
        if system == "Windows":
            return self._install_windows(font_file_path, target_filename)
        elif system == "Linux":
            return self._install_linux(font_file_path, target_filename)
        elif system == "Darwin":
            return self._install_mac(font_file_path, target_filename)
        return False, f"不支持的操作系统: {system}"

    # ------------------------------------------------------------------
    # 主流程
    # ------------------------------------------------------------------
    def ensure_fonts_available(self) -> Dict[str, Dict]:
        """依次检测必需字体，缺失时直接使用内置字体文件安装。"""
        result = {}

        for font_name in self._required_fonts().keys():
            installed = self.check_font_installed(font_name)
            item = {
                'name': font_name,
                'installed': installed,
                'action_taken': None,
                'success': installed,
                'message': '',
            }

            if installed:
                item['message'] = f"字体 {font_name} 已安装"
                result[font_name] = item
                continue

            print(f"\n检测到字体 {font_name} 未安装，开始安装...")

            bundled_path = self._get_bundled_font_path(font_name)
            if not bundled_path:
                item['action_taken'] = 'miss_bundled'
                item['message'] = f"缺少内置字体文件: {font_name}，请手动下载"
                result[font_name] = item
                continue

            target_filename = self._get_target_filename(font_name)
            ok, msg = self.install_font_file(bundled_path, target_filename)

            item['action_taken'] = 'installed'
            item['success'] = ok
            item['message'] = msg

            if ok:
                # 安装后强制刷新缓存，保证后续检测准确
                self._installed_families_cache = None
            result[font_name] = item

        return result

    # ------------------------------------------------------------------
    # 手动安装指引（兜底）
    # ------------------------------------------------------------------
    def get_manual_install_guide(self, font_name: str) -> str:
        info = self._required_fonts().get(font_name)
        if not info:
            return f"字体 {font_name} 不在配置列表中"

        bundled = self._get_bundled_font_path(font_name)
        sources = info.get("sources", [])

        guide = f"\n{'=' * 60}\n"
        guide += f"字体 {font_name} 手动安装指南\n"
        guide += f"{'=' * 60}\n\n"

        if bundled:
            guide += f"内置字体文件已就位: {bundled}\n\n"

        if sources:
            guide += "网上参考下载地址:\n"
            for i, source in enumerate(sources, 1):
                guide += f"{i}. {source['name']}: {source['url']}\n"

        guide += "\n安装步骤:\n"
        system = platform.system()

        if system == "Windows":
            guide += "1. 双击打开字体文件（.ttf），点击“安装”\n"
            guide += "2. 或将其复制到 C:\\Windows\\Fonts 目录\n"
            guide += "3. 重启目标应用程序使字体生效\n"
        elif system == "Darwin":
            guide += "1. 双击字体文件，点击“安装字体”\n"
            guide += "   或复制到 ~/Library/Fonts 目录\n"
        elif system == "Linux":
            guide += "1. 复制字体到 ~/.fonts 目录\n"
            guide += "2. 运行: fc-cache -f\n"

        guide += f"\n{'=' * 60}\n"
        return guide


def main():
    print("公文格式化 - 字体管理工具")
    print("=" * 60)

    manager = FontManager()

    print("\n检查系统字体安装状态...")
    fonts_status = manager.check_all_fonts()

    all_installed = all(fonts_status.values())

    if all_installed:
        print("\n✓ 所有必需字体已安装:")
        for font_name, installed in fonts_status.items():
            print(f"  - {font_name}: ✓")
    else:
        print("\n✗ 以下字体未安装:")
        for font_name, installed in fonts_status.items():
            status = "✓" if installed else "✗"
            print(f"  - {font_name}: {status}")

        print("\n开始自动安装字体...")
        result = manager.ensure_fonts_available()

        print("\n安装结果:")
        for font_name, info in result.items():
            status = "✓" if info['success'] else "✗"
            print(f"  {status} {font_name}: {info['message']}")
            if not info['success']:
                print(manager.get_manual_install_guide(font_name))

    print("\n字体检查完成！")


if __name__ == "__main__":
    main()