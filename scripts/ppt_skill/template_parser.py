"""
PPT模板解析器 - 从 .pptx 模板中提取配色、字体、版式等风格参数。

用法:
    python template_parser.py --input template.pptx
"""

import json
import sys
import zipfile
from pathlib import Path

from lxml import etree


NS = {'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'}


def parse_theme(pptx_path: str) -> dict:
    """从 .pptx 文件中提取主题样式参数。

    返回:
        {
            "colors": {
                "primary": "#4F81BD",
                "secondary": "#C0504D",
                "accent1": "#4F81BD",
                "accent2": "#C0504D",
                "accent3": "#9BBB59",
                "accent4": "#8064A2",
                "accent5": "#4BACC6",
                "accent6": "#F79646",
                "dark1": "#000000",
                "light1": "#FFFFFF",
                "dark2": "#1F497D",
                "light2": "#EEECE1"
            },
            "fonts": {
                "title_latin": "Calibri",
                "body_latin": "Calibri",
                "title_ea": "微软雅黑",
                "body_ea": "微软雅黑"
            },
            "slide_size": {"width": 13.333, "height": 7.5},
            "fill_styles": ["solid", "gradient", "gradient"]
        }
    """
    result = {
        "colors": {},
        "fonts": {},
        "slide_size": {"width": 13.333, "height": 7.5},
        "fill_styles": [],
    }

    # 从主题文件读取
    with zipfile.ZipFile(pptx_path, 'r') as z:
        # 找主题文件
        theme_files = [f for f in z.namelist() if f.startswith('ppt/theme/') and f.endswith('.xml')]
        if not theme_files:
            return result

        theme_xml = z.read(theme_files[0])
        root = etree.fromstring(theme_xml)

        # --- 主题色 ---
        clr_scheme = root.find('.//a:clrScheme', NS)
        if clr_scheme is not None:
            for child in clr_scheme:
                tag = child.tag.split('}')[-1]
                srgb = child.find('a:srgbClr', NS)
                if srgb is not None:
                    result["colors"][tag] = f"#{srgb.get('val')}"
                else:
                    sys_clr = child.find('a:sysClr', NS)
                    if sys_clr is not None:
                        # 系统颜色，映射常用值
                        val = sys_clr.get('val')
                        mapping = {
                            'windowText': '#000000',
                            'window': '#FFFFFF',
                        }
                        result["colors"][tag] = mapping.get(val, val)

        # --- 字体 ---
        font_scheme = root.find('.//a:fontScheme', NS)
        if font_scheme is not None:
            for ftype, label in [('a:majorFont', 'title'), ('a:minorFont', 'body')]:
                fonts = font_scheme.find(ftype, NS)
                if fonts is not None:
                    for lang_attr, lang_label in [('a:latin', 'latin'), ('a:ea', 'ea'), ('a:cs', 'cs')]:
                        f = fonts.find(lang_attr, NS)
                        if f is not None:
                            result["fonts"][f"{label}_{lang_label}"] = f.get('typeface', '')

        # --- 填充样式 ---
        fmt_scheme = root.find('.//a:fmtScheme', NS)
        if fmt_scheme is not None:
            fill_style_lst = fmt_scheme.find('a:fillStyleLst', NS)
            if fill_style_lst is not None:
                for fill in fill_style_lst:
                    tag = fill.tag.split('}')[-1]
                    result["fill_styles"].append(tag)

    # 从 PPT 文件读取页面尺寸
    try:
        with zipfile.ZipFile(pptx_path, 'r') as z:
            pres_xml = z.read('ppt/presentation.xml')
            pres_root = etree.fromstring(pres_xml)
            p_ns = {'p': 'http://schemas.openxmlformats.org/presentationml/2006/main'}
            sld_sz = pres_root.find('.//p:sldSz', p_ns)
            if sld_sz is not None:
                cx = int(sld_sz.get('cx', 0)) / 914400  # EMU to inches
                cy = int(sld_sz.get('cy', 0)) / 914400
                result["slide_size"] = {"width": round(cx, 3), "height": round(cy, 3)}
    except Exception:
        pass

    return result


def format_as_params(theme: dict) -> dict:
    """将主题参数转换为 CLI 可用格式。"""
    colors = theme.get("colors", {})
    fonts = theme.get("fonts", {})

    params = {
        "primary_color": colors.get("accent1", "#4F81BD"),
        "secondary_color": colors.get("accent2", "#C0504D"),
        "accent_color": colors.get("accent6", "#F79646"),
        "background_color": colors.get("light1", "#FFFFFF"),
        "title_font": fonts.get("title_latin", "Calibri"),
        "body_font": fonts.get("body_latin", "Calibri"),
        "ea_title_font": fonts.get("title_ea", ""),
        "ea_body_font": fonts.get("body_ea", ""),
        "slide_width": theme.get("slide_size", {}).get("width", 13.333),
        "slide_height": theme.get("slide_size", {}).get("height", 7.5),
    }
    return params


def main():
    if len(sys.argv) != 3 or sys.argv[1] != '--input':
        print("用法: python template_parser.py --input <template.pptx>")
        sys.exit(1)

    pptx_path = sys.argv[2]
    if not Path(pptx_path).exists():
        print(f"错误: 文件 {pptx_path} 不存在")
        sys.exit(1)

    theme = parse_theme(pptx_path)
    params = format_as_params(theme)

    print(json.dumps(params, indent=2, ensure_ascii=False))


if __name__ == '__main__':
    main()
