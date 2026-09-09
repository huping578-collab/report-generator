from __future__ import annotations

import base64
import io
import json
import os
import openpyxl
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from unittest.mock import patch

from docx import Document
from docx.oxml.ns import qn

from bridge import DesktopBridge
from backend import markdown_skeleton, minimal_docx, report_engine as engine


def test_desktop_apple_style_layout():
    """Render real HTML/JS in Chrome; no desktop API or report generation is mocked."""
    import html
    import re
    import subprocess

    chrome = Path('C:/Program Files/Google/Chrome/Application/chrome.exe')
    assert chrome.is_file(), 'Chrome is required for the desktop layout check'
    frontend = Path(__file__).resolve().parents[1] / 'frontend'
    artifacts = Path('D:/hermes/tests/report-generator-ui')
    artifacts.mkdir(parents=True, exist_ok=True)
    source = (frontend / 'index.html').read_text(encoding='utf-8')
    script = (frontend / 'app.js').read_text(encoding='utf-8')
    probe = r'''<script>
    const errors = [];
    window.addEventListener('error', e => errors.push(e.message));
    setTimeout(() => {
      const results = [];
      for (const template of ['cq', 'gd']) {
        document.querySelector(`[data-template="${template}"]`).click();
        const visible = [...document.querySelectorAll('.path-field')]
          .filter(e => e.getBoundingClientRect().width > 0);
        results.push({template, title: document.querySelector('#pageTitle').textContent,
          overflow: document.documentElement.scrollWidth > innerWidth,
          narrow: visible.filter(e => e.getBoundingClientRect().width < 30).map(e => e.id),
          hidden: [...document.querySelectorAll(template === 'cq' ? '.only-gd' : '.only-cq')]
            .every(e => e.getBoundingClientRect().width === 0)});
      }
      document.querySelector('[data-template="cq"]').click();
      const out = document.createElement('pre'); out.id = 'ui-check'; out.hidden = true;
      out.textContent = JSON.stringify({results, errors,
        blur: getComputedStyle(document.querySelector('.app-shell')).backdropFilter,
        warning: getComputedStyle(document.querySelector('#configStatus')).color});
      document.body.append(out);
    }, 1000);
    </script>'''
    page = artifacts / 'render-check.html'
    page.write_text(source.replace('<script src="app.js"></script>',
                                  '<script>' + script + '</script>' + probe), encoding='utf-8')
    for width, height in [(1380, 880), (1121, 800), (920, 680), (390, 844)]:
        with tempfile.TemporaryDirectory(dir=artifacts) as profile:
            result = subprocess.run([str(chrome), '--headless=new', '--no-first-run',
                '--no-default-browser-check', '--hide-scrollbars',
                '--force-device-scale-factor=1', f'--window-size={width},{height}',
                '--virtual-time-budget=2000', f'--user-data-dir={profile}',
                f'--screenshot={artifacts / f"layout-{width}.png"}', '--dump-dom', page.as_uri()],
                capture_output=True, encoding='utf-8', errors='replace', timeout=60)
        assert result.returncode == 0, result.stderr
        match = re.search(r'<pre id="ui-check"[^>]*>(.*?)</pre>', result.stdout, re.S)
        assert match, result.stderr
        data = json.loads(html.unescape(match.group(1)))
        (artifacts / f'layout-{width}.json').write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
        assert not data['errors'], data
        assert 'blur(24px)' in data['blur'], data
        for row in data['results']:
            assert not row['overflow'] and not row['narrow'] and row['hidden'], (width, row)
            assert ('重庆' if row['template'] == 'cq' else '广东') in row['title']


def test_chongqing_tci_units_keep_original_boundaries_and_equal_weights():
    import pytest
    segments = [
        dict(county="甲县", route="G210", start=2157392., end=2159964., mileage=2.572),
        dict(county="甲县", route="G210", start=2159964., end=2160200., mileage=.236),
        dict(county="甲县", route="G210", start=2161500., end=2161700., mileage=.2),
        dict(county="乙县", route="G210", start=2157392., end=2159964., mileage=2.572),
    ]
    records = [dict(segment=0, direction="上行", station=2157500., light=20, heavy=0, sign=0, marking=0),
               dict(segment=0, direction="下行", station=2158000., light=0, heavy=0, sign=1, marking=11),
               dict(segment=1, direction="上行", station=2159964., light=1, heavy=0, sign=0, marking=0)]
    stats = engine.make_tci_stats(segments, records)
    units = [u for s in stats for u in s.get("units", [])]
    assert len(units) == 8
    assert [(u["start"], u["end"]) for u in units if u["direction"] == "上行"] == [
        (2157392., 2158000.), (2158000., 2159000.), (2159000., 2159964.),
        (2159964., 2160000.), (2160000., 2160200.)]
    assert stats[2]["tci"] is None and stats[3]["tci"] is None
    assert stats[0]["units"][1]["tci"] == 100
    groups = engine.tci_route_stats(segments, stats)
    up = next(g for g in groups if g["county"] == "甲县" and g["direction"] == "上行")
    expected = (engine.compute_tci(20, 0, 0, 0) + 100 + 100 + engine.compute_tci(1, 0, 0, 0) + 100) / 5
    assert up["tci"] == pytest.approx(expected)
    assert up["mileage"] == pytest.approx(2.808)
    assert up["grade"] == engine.tci_grade(expected)


def test_chongqing_tci_source_identity_and_boundary_assignment():
    root = Path(__file__).parent / "artifacts" / "tci-identity"
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county=c, route=r, start=a, end=b, direction=d)
                for c, r, a, b, d in [("甲县", "G210", 1000, 1964, "上行"),
                                      ("甲县", "G210", 1964, 3000, "上行"),
                                      ("甲县", "G210", 1000, 3000, "下行"),
                                      ("乙县", "G210", 1000, 3000, "上行"),
                                      ("甲县", "G319", 1000, 3000, "上行")]]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["区域", "路线编号", "方向", "原始桩号", "防护设施缺损", "标志缺损", "标线缺损"])
    ws.append([None, None, None, None, "轻", None, None])
    for c, r, d, station in [("甲县", "G210", "上行", "K1+964"), ("甲县", "G210", "下行", "K1+500"),
                              ("乙县", "G210", "上行", "K1+500"), ("甲县", "G319", "上行", "K1+500")]:
        ws.append([c, r, d, station, 1, 0, 0])
    path = root / "identity.xlsx"
    wb.save(path)
    records = engine.collect_tci_records(segments, path)
    assert [r["segment"] for r in records] == [1, 2, 3, 4]
    assert [r["direction"] for r in records] == ["上行", "下行", "上行", "上行"]
    assert records[2]["county"] == "乙县"
    # 市域标签不是区县；单县输入仍应落段。
    ws.delete_rows(3, 4)
    ws.append(["重庆市", "G210", "上行", "K1+500", 1, 0, 0])
    wb.save(path)
    assert len(engine.collect_tci_records(segments[:2], path)) == 1


def test_chongqing_route_report_and_appendix_end_to_end():
    root = Path(__file__).parent / "artifacts" / "chongqing-route"
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county="甲县", route=route, start=start, end=end, mileage=(end-start)/1000,
                     grade="二级公路", manager="", route_name="测试路线")
                for route, start, end in [("G210", 1500, 2500), ("G210", 3964, 5000), ("G319", 6000, 7000)]]
    h = [dict(segment=i, kind="二波", height=value, direction="上行", station=s["start"]+10,
              raw_station=s["start"]+10, electronic_station=s["start"]+10, file="test.xlsx", basis="电子修正桩号")
         for i, s in enumerate(segments) for value in (600, 500)]
    b = [dict(segment=i, direction="上行", station=s["start"]+10, raw_station=s["start"]+10,
              file="test.xlsx", basis="电子修正桩号", splice=10, connection=5, splice_missing=1, connection_missing=0)
         for i, s in enumerate(segments)]
    t = [dict(segment=i, direction="上行", station=s["start"]+10, light=0, heavy=0, sign=1, marking=0)
         for i, s in enumerate(segments)]
    stats = engine.make_tci_stats(segments, t)
    config = engine.Config(root, root/'summary.xlsx', root, engine.builtin_template_paths()["重庆模板"], root, county="甲县")
    minimal_docx.make_report(config, segments, engine.make_stats(segments,h), h,
                            engine.make_bolt_stats(segments,b), b, stats, t, {}, root,
                            skeleton_md=config.template_docx)
    engine.make_excel(config, segments, tci_stats=stats, tci_records=t)
    doc = Document(config.out_docx)
    headings = [p.text for p in doc.paragraphs if p.style.name == 'Heading 2']
    assert not any('线K' in text for text in headings)
    assert sum(text.endswith('G210线') for text in headings) == 3
    assert sum(text.endswith('G319线') for text in headings) == 3
    text = '\n'.join(p.text for p in doc.paragraphs)
    assert '附表1 重庆市甲县交安设施技术状况评定明细' in text
    assert '评定单元等权' in text
    assert '共检出拼接螺栓20颗' in text
    appendix = next(table for table in doc.tables if [c.text for c in table.rows[0].cells] ==
                    ['序号','路线编号','区县','方向','起点桩号','止点桩号','TCI','等级'])
    assert len(appendix.rows) == 8  # 5 个单元 + 2 条路线汇总 + 表头
    assert appendix.rows[0]._tr.xpath('./w:trPr/w:tblHeader')
    assert not doc._element.xpath('.//w:r/w:r')  # TOC缓存不得生成嵌套run。
    assert sum("附表1 重庆市甲县交安设施技术状况评定明细" in p.text for p in doc.paragraphs) == 2
    assert len(doc.inline_shapes) >= 6
    wb = openpyxl.load_workbook(config.out_xlsx, read_only=True, data_only=True)
    assert wb['TCI公里评定明细'].max_row == 6
    wb.close()


def test_chongqing_real_tci_workbook_matches_reference_units():
    import pytest
    source = Path(os.environ.get("CHONGQING_TCI_TEST_SOURCE", "D:/hermes/attachments/G210上行2159.964--2184.977.xlsx"))
    if not source.is_file():
        pytest.skip("未配置真实TCI验算工作簿")
    root = Path(__file__).parent / "artifacts" / "chongqing-route" / "real-tci"
    root.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    expected = [r for r in wb['技术状况评定表'].iter_rows(min_row=2,values_only=True)
                if isinstance(r[1], (int,float)) and isinstance(r[2], (int,float)) and isinstance(r[7], (int,float))]
    wb.close()
    # 检测范围取真实评定表，不从病害点的最小/最大桩号猜测覆盖范围。
    start, end = expected[0][1], expected[-1][2]
    summary = openpyxl.Workbook()
    sheet = summary.active
    sheet.title = '各区县项目概况'
    sheet.append(['区县','路线编号','公路等级','起点桩号','止点桩号','里程（km）'])
    sheet.append(['两江新区','G210','一级公路',start,end,end-start])
    summary_path = root/'summary.xlsx'
    summary.save(summary_path)
    config = engine.Config(root, summary_path, root, engine.builtin_template_paths()['重庆模板'], root, tci_path=source, county='两江新区')
    result = engine.generate_statistics_and_report(config, process_height=False, process_tci=True, county_override='两江新区')
    units = result['tci'][0]['units']
    assert len(units) == len(expected) == 26
    for unit, reference in zip(units,expected):
        assert unit['start'] == pytest.approx(reference[1]*1000)
        assert unit['end'] == pytest.approx(reference[2]*1000)
        assert unit['tci'] == pytest.approx(reference[7], abs=.0001)
        assert unit['grade'] == reference[8]
    assert result['tci'][0]['tci'] == pytest.approx(87.50549450549453, abs=.0001)
    doc = Document(config.out_docx)
    assert any('87.51' in p.text for p in doc.paragraphs)
    appendix = next(t for t in doc.tables if [c.text for c in t.rows[0].cells] ==
                    ['序号','路线编号','区县','方向','起点桩号','止点桩号','TCI','等级'])
    assert len(appendix.rows) == 28
    assert appendix.rows[1].cells[6].text == '87.51'
    with zipfile.ZipFile(config.out_docx) as archive:
        assert archive.testzip() is None


def test_tci_charts_split_by_original_segment_and_direction():
    from PIL import Image
    root = Path(__file__).parent / 'artifacts' / 'chongqing-refinement'
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county='甲县',route='G210',start=a,end=b,mileage=(b-a)/1000)
                for a,b in [(1500,3500),(5000,6500)]]
    rows = [dict(segment=i,direction=d,station=segments[i]['start']+20,light=0,heavy=0,sign=1,marking=0)
            for i,d in [(0,'上行'),(0,'下行'),(1,'上行')]]
    stats = engine.make_tci_stats(segments,rows)
    figures = []
    subplots = engine.plt.subplots
    def observe(*args,**kwargs):
        pair = subplots(*args,**kwargs)
        figures.append(pair)
        return pair
    with patch.object(engine.plt,'subplots',side_effect=observe):
        images = engine.report_tci_images(root,segments,stats)
    assert set(images) == {(0,'上行'),(0,'下行'),(1,'上行')}
    for (figure,axis),path in zip(figures,images.values()):
        assert tuple(round(v,5) for v in figure.get_size_inches()) == (round(13/2.54,5),round(8/2.54,5))
        assert figure.dpi == 180
        assert axis.get_title() == ''
        assert axis.lines[0].get_color() == '#4472C4'
        assert axis.lines[0].get_linewidth() == 1
        assert not axis.get_legend().get_frame_on()
        assert axis.get_xticklabels()[0].get_rotation() == 30
        with Image.open(path) as image:
            assert image.width > image.height
    doc = Document()
    minimal_docx._section_tci(doc,segments,stats,images,root)
    assert len(doc.inline_shapes) == 3
    assert [p.text for p in doc.paragraphs if p.style.name == 'Heading 2'] == ['甲县整体情况','G210线']
    captions = [p.text for p in doc.paragraphs if 'TCI情况' in p.text]
    assert any('上行K1+500～K3+500段TCI情况' in text for text in captions)
    assert any('下行K1+500～K3+500段TCI情况' in text for text in captions)


def test_conclusion_matches_reference_template_structure():
    root = Path(__file__).parent / 'artifacts' / 'chongqing-refinement'
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county='甲县', route='G210', start=1500, end=2500, mileage=1.0,
                     grade='一级公路', manager='', route_name='测试')]
    h = [dict(segment=0, kind='二波', height=500, direction='上行', station=1510,
              raw_station=1510, electronic_station=1510, file='t.xlsx', basis='电子修正桩号')]
    b = [dict(segment=0, direction='上行', station=1510, raw_station=1510,
              file='t.xlsx', basis='电子修正桩号', splice=100, connection=50, splice_missing=5, connection_missing=0)]
    t = [dict(segment=0, direction='上行', station=1510, light=1, heavy=0, sign=2, marking=10.0),
         dict(segment=0, direction='下行', station=1520, light=10, heavy=5, sign=10, marking=100.0)]
    stats = engine.make_tci_stats(segments, t)
    doc = Document()
    minimal_docx._section_conclusion(doc, segments, engine.make_stats(segments, h),
                                     engine.make_bolt_stats(segments, b), stats)
    text = '\n'.join(p.text for p in doc.paragraphs)
    headings = [p.text for p in doc.paragraphs if p.style.name.startswith('Heading')]
    assert '总结' in headings and '建议' in headings
    assert '沿线设施技术状况' in text
    assert '合格率' in text and '螺栓' in text and '缺失率' in text
    assert 'JTG 5110-2023' in text
    assert '1.沿线设施技术状况检查建议' in text
    assert '2.波形梁护栏专项检测建议' in text
    assert '标志遮挡' in text and '标线缺损' in text and '波形梁变形' in text
    assert '横梁中心高度' in text and '螺栓缺失' in text


def test_tci_chart_thins_station_ticks_when_many_points():
    root = Path(__file__).parent / 'artifacts' / 'chongqing-refinement'
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county='甲县', route='G210', start=1000, end=27000, mileage=26.0)]
    rows = [dict(segment=0, direction='上行', station=1000 + i * 1000 + 20, light=0, heavy=0, sign=1, marking=0)
            for i in range(26)]
    stats = engine.make_tci_stats(segments, rows)
    captured = {}
    subplots = engine.plt.subplots
    def observe(*args, **kwargs):
        pair = subplots(*args, **kwargs)
        captured['axis'] = pair[1]
        return pair
    with patch.object(engine.plt, 'subplots', side_effect=observe):
        engine.report_tci_images(root, segments, stats)
    labels = captured['axis'].get_xticklabels()
    assert 2 <= len(labels) <= 10
    assert labels[0].get_text() == 'K1+000'


def test_conclusion_lists_worst_five_weak_segments_per_route():
    segments = [dict(county='甲县', route='G210', start=1000 + i * 1000, end=2000 + i * 1000,
                     mileage=1.0, grade='一级公路', manager='', route_name='测试') for i in range(6)]
    stats = [dict(units=[dict(route='G210', county='甲县', direction='上行',
                               start=seg['start'], end=seg['end'], tci=60.0 + i)])
             for i, seg in enumerate(segments)]
    doc = Document()
    minimal_docx._section_conclusion(doc, segments, None, None, stats)
    sentence = next(p.text for p in doc.paragraphs if p.text.startswith('甲县沿线设施技术状况TCI整体状况'))
    assert sentence == ('甲县沿线设施技术状况TCI整体状况一般，其中G210线K1+000～K2+000段、K2+000～K3+000段、'
                        'K3+000～K4+000段、K4+000～K5+000段、K5+000～K6+000段等评定为次差，需要特别注意。')


def test_conclusion_groups_weak_segments_by_route():
    segments = [dict(county='甲县', route=route, start=start, end=start + 1000, mileage=1.0,
                     grade='一级公路', manager='', route_name='测试')
                for route, start in (('G210', 1000), ('G319', 3000))]
    stats = [dict(units=[dict(route=route, county='甲县', direction='上行',
                               start=start, end=start + 1000, tci=tci)])
             for (route, start), tci in zip((('G210', 1000), ('G319', 3000)), (65.0, 68.0))]
    doc = Document()
    minimal_docx._section_conclusion(doc, segments, None, None, stats)
    sentence = next(p.text for p in doc.paragraphs if p.text.startswith('甲县沿线设施技术状况TCI整体状况'))
    assert sentence == ('甲县沿线设施技术状况TCI整体状况一般，其中G210线K1+000～K2+000段；'
                        'G319线K3+000～K4+000段评定为次差，需要特别注意。')


def test_conclusion_merges_same_range_from_both_directions():
    segments = [dict(county='甲县', route='G210', start=1000, end=2000, mileage=1.0,
                     grade='一级公路', manager='', route_name='测试')]
    stats = [dict(units=[dict(route='G210', county='甲县', direction=direction, start=1000, end=2000, tci=tci)
                         for direction, tci in (('上行', 65.0), ('下行', 66.0))])]
    doc = Document()
    minimal_docx._section_conclusion(doc, segments, None, None, stats)
    sentence = next(p.text for p in doc.paragraphs if p.text.startswith('甲县沿线设施技术状况TCI整体状况'))
    assert sentence == '甲县沿线设施技术状况TCI整体状况一般，其中G210线K1+000～K2+000段评定为次差，需要特别注意。'


def test_example_tables_are_borderless():
    from docx import Document as Doc
    from docx.oxml.ns import qn as qqn
    root = Path(__file__).parent / 'artifacts' / 'chongqing-refinement'
    root.mkdir(parents=True, exist_ok=True)
    segments = [dict(county='甲县', route='G210', start=1500, end=2500, mileage=1.0,
                     grade='一级公路', manager='', route_name='测试')]
    h = [dict(segment=0, kind='二波', height=600, direction='上行', station=1510,
              raw_station=1510, electronic_station=1510, file='t.xlsx', basis='电子修正桩号')]
    config = engine.Config(root, root/'s.xlsx', root, engine.builtin_template_paths()['重庆模板'], root, county='甲县')
    minimal_docx.make_report(config, segments, engine.make_stats(segments,h), h,
                            engine.make_bolt_stats(segments,[]), [], None, None, {}, root,
                            skeleton_md=config.template_docx)
    doc = Doc(config.out_docx)
    for table in doc.tables:
        if len(table.columns) <= 2:
            borders = table._tbl.tblPr.find(qqn('w:tblBorders'))
            if borders is not None:
                for edge in ('top','left','bottom','right','insideH','insideV'):
                    el = borders.find(qqn(f'w:{edge}'))
                    assert el is not None and el.get(qqn('w:val')) == 'none', f'table has visible border {edge}'


class FakeWindow:
    def __init__(self) -> None:
        self.scripts: list[str] = []

    def evaluate_js(self, script: str) -> None:
        self.scripts.append(script)


class MarkdownSkeletonTests(unittest.TestCase):
    def test_template_reads_toml_front_matter_and_body_blocks(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "template.md"
            path.write_text(
                "+++\n"
                "[page]\n"
                "orientation = \"landscape\"\n"
                "[body]\n"
                "size_pt = 11\n"
                "+++\n\n"
                "# 标题\n\n"
                "<!-- toc -->\n\n"
                "正文。\n",
                encoding="utf-8",
            )

            template = markdown_skeleton.read_template(path)
            self.assertEqual(markdown_skeleton.read_blocks(path), template.blocks)

        self.assertEqual(template.config["page"]["orientation"], "landscape")
        self.assertEqual(template.config["body"]["size_pt"], 11)
        self.assertEqual([block.kind for block in template.blocks], ["heading", "toc", "paragraph"])
        self.assertEqual(template.blocks[-1].text, "正文。")

    def test_template_without_front_matter_uses_legacy_defaults(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "legacy.md"
            path.write_text("# 旧模板\n\n正文。\n", encoding="utf-8")

            template = markdown_skeleton.read_template(path)
            self.assertEqual(
                markdown_skeleton.read_blocks(path),
                [
                    markdown_skeleton.Block("heading", level=1, text="旧模板"),
                    markdown_skeleton.Block("paragraph", text="正文。"),
                ],
            )

        self.assertEqual(template.config["page"]["paper"], "A4")
        self.assertEqual(template.config["page"]["orientation"], "portrait")
        self.assertEqual(template.config["body"]["east_asia"], "仿宋_GB2312")
        self.assertEqual(template.config["body"]["latin"], "Times New Roman")
        self.assertEqual(template.config["body"]["size_pt"], 10.5)
        self.assertEqual(template.config["body"]["first_line_chars"], 2)
        self.assertFalse(template.config["toc"]["enabled"])

    def test_default_config_isolated_between_template_reads(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "legacy.md"
            path.write_text("# 旧模板\n", encoding="utf-8")

            first = markdown_skeleton.read_template(path)
            first.config["body"]["size_pt"] = 99
            second = markdown_skeleton.read_template(path)

        self.assertEqual(second.config["body"]["size_pt"], 10.5)
        self.assertEqual(markdown_skeleton.DEFAULT_CONFIG["body"]["size_pt"], 10.5)

    def test_toml_multiline_string_can_contain_separator_line(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "multiline.md"
            path.write_text(
                '''+++
[toc]
enabled = false
title = """目录
+++
说明"""
+++
# 标题
''',
                encoding="utf-8",
            )

            template = markdown_skeleton.read_template(path)

        self.assertEqual(template.config["toc"]["title"], "目录\n+++\n说明")
        self.assertEqual([block.text for block in template.blocks], ["标题"])

    def test_repository_templates_declare_format_and_toc_policy(self) -> None:
        root = Path(__file__).resolve().parents[1]
        chongqing = markdown_skeleton.read_template(root / "templates" / "重庆项目报告模板.md")
        guangdong = markdown_skeleton.read_template(root / "templates" / "广东项目第五章模板.md")

        self.assertTrue(chongqing.config["toc"]["enabled"])
        self.assertTrue(any(block.kind == "toc" for block in chongqing.blocks))
        self.assertFalse(guangdong.config["toc"]["enabled"])
        self.assertFalse(any(block.kind == "toc" for block in guangdong.blocks))
        for config in (chongqing.config, guangdong.config):
            self.assertEqual(set(config), {"page", "body", "heading", "table", "caption", "toc"})

    def test_chongqing_template_conclusion_has_single_injection_without_static_duplicates(self) -> None:
        # D2 回归锁：结论以程序注入为准，模板不得残留静态 5.1结论/5.2建议。
        from backend import minimal_docx

        root = Path(__file__).resolve().parents[1]
        template = markdown_skeleton.read_template(root / "templates" / "重庆项目报告模板.md")
        anchors = [block for block in template.blocks if block.text and "inject:conclusion" in block.text]
        self.assertEqual(len(anchors), 1)
        static_leftovers = [
            block.text
            for block in template.blocks
            if block.kind == "heading" and minimal_docx._strip_section_number(block.text) in ("结论", "建议")
        ]
        self.assertEqual(static_leftovers, [])

    def test_invalid_template_config_reports_path_and_field(self) -> None:
        cases = (
            ("[body]\nsize_pt = 73\n", "size_pt"),
            ("[page]\nleft_margin_cm = 11\n", "left_margin_cm"),
            ("[page]\norientation = \"diagonal\"\n", "orientation"),
        )
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "invalid.md"
            for front_matter, field in cases:
                path.write_text(f"+++\n{front_matter}+++\n# 标题\n", encoding="utf-8")
                with self.subTest(field=field):
                    with self.assertRaises(ValueError) as raised:
                        markdown_skeleton.read_template(path)
                    self.assertIn(str(path), str(raised.exception))
                    self.assertIn(field, str(raised.exception))

    def test_toc_enabled_requires_toc_marker_but_disabled_allows_legacy_body(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "重庆模板.md"
            path.write_text("+++\n[toc]\nenabled = true\n+++\n# 标题\n", encoding="utf-8")
            with self.assertRaises(ValueError) as raised:
                markdown_skeleton.read_template(path)
            self.assertIn(str(path), str(raised.exception))
            self.assertIn("toc", str(raised.exception))

            path.write_text("+++\n[toc]\nenabled = false\n+++\n# 广东模板\n", encoding="utf-8")
            template = markdown_skeleton.read_template(path)

        self.assertEqual([block.kind for block in template.blocks], ["heading"])
        self.assertFalse(template.config["toc"]["enabled"])


class MarkdownDocxFormattingTests(unittest.TestCase):
    def _write_report(self, template_text: str) -> Path:
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        root = Path(temp_dir.name)
        template = root / "template.md"
        template.write_text(template_text, encoding="utf-8")
        output = root / "output"
        output.mkdir()
        report_config = engine.Config(
            project_dir=root,
            summary_xlsx=root / "summary.xlsx",
            detail_dir=root,
            template_docx=template,
            output_dir=output,
        )
        minimal_docx.make_report(
            report_config, [], None, [], None, [], None, [], {}, root,
            skeleton_md=template,
        )
        return output / engine.OUT_DOCX_NAME

    def test_markdown_config_controls_docx_runs_and_outline(self) -> None:
        template_text = (
            "+++\n"
            "[body]\n"
            "east_asia = \"楷体\"\n"
            "latin = \"Arial\"\n"
            "size_pt = 11\n"
            "first_line_chars = 0\n"
            "[heading.1]\n"
            "east_asia = \"黑体\"\n"
            "latin = \"Arial\"\n"
            "size_pt = 18\n"
            "outline_level = 0\n"
            "[table]\n"
            "east_asia = \"宋体\"\n"
            "latin = \"Arial\"\n"
            "size_pt = 9\n"
            "[caption]\n"
            "east_asia = \"黑体\"\n"
            "latin = \"Arial\"\n"
            "size_pt = 12\n"
            "[toc]\n"
            "enabled = true\n"
            "update_on_open = true\n"
            "+++\n"
            "# 自定义标题\n\n"
            "<!-- toc -->\n\n"
            "<!-- inject:overview -->\n"
        )
        report_path = self._write_report(template_text)
        with zipfile.ZipFile(report_path) as archive:
            document = archive.read("word/document.xml").decode("utf-8")
            settings = archive.read("word/settings.xml").decode("utf-8")
        self.assertIn('w:eastAsia="黑体"', document)
        self.assertIn('w:eastAsia="楷体"', document)
        self.assertIn('w:eastAsia="宋体"', document)
        self.assertIn('w:ascii="Arial"', document)
        self.assertIn('w:hAnsi="Arial"', document)
        self.assertIn('w:val="0"', document)
        self.assertIn('w:instrText xml:space="preserve"> TOC \\o "1-5" \\h \\z \\u </w:instrText>', document)
        self.assertIn('w:updateFields w:val="true"', settings)

    def test_disabled_toc_does_not_write_toc_field(self) -> None:
        report_path = self._write_report("+++\n[toc]\nenabled = false\n+++\n# 无目录\n")
        with zipfile.ZipFile(report_path) as archive:
            document = archive.read("word/document.xml").decode("utf-8")
        self.assertNotIn(" TOC ", document)


class GuangdongTemplateConfigTests(unittest.TestCase):
    @staticmethod
    def _bundle() -> dict:
        return {
            "marking": [],
            "height": [],
            "bolt": [],
            "notes": [],
            "comparison_detail": [],
            "weak_segments": [],
        }

    @staticmethod
    def _write_report(root: Path, template: Path, output_dir: Path | None = None) -> Path:
        return engine.GuangdongChapterWriter.write(
            "佛山市",
            GuangdongTemplateConfigTests._bundle(),
            output_dir or root / "output",
            template,
            {"marking": 7, "height": 5, "bolt": 5},
        )

    @staticmethod
    def _xml_paragraphs(document_xml: bytes):
        from xml.etree import ElementTree as ET

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(document_xml)
        return root.findall(".//w:body/w:p", ns), ns

    def test_guangdong_writer_disables_toc_and_update_fields_from_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "disabled-toc.md"
            template.write_text(
                "+++\n"
                "[toc]\n"
                "enabled = false\n"
                "+++\n"
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "<!-- toc -->\n",
                encoding="utf-8",
            )
            output = self._write_report(root, template)
            with zipfile.ZipFile(output) as archive:
                document = archive.read("word/document.xml").decode("utf-8")
                settings = archive.read("word/settings.xml").decode("utf-8")

        self.assertNotIn(" TOC ", document)
        self.assertNotIn("w:updateFields", settings)

    def test_guangdong_writer_uses_custom_toc_range_title_and_update_policy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "custom-toc.md"
            template.write_text(
                "+++\n"
                "[heading.1]\n"
                "east_asia = \"目录标题字体\"\n"
                "size_pt = 18\n"
                "[toc]\n"
                "enabled = true\n"
                "title = \"自定义目录\"\n"
                "min_level = 2\n"
                "max_level = 4\n"
                "update_on_open = false\n"
                "+++\n"
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "<!-- toc -->\n",
                encoding="utf-8",
            )
            output = self._write_report(root, template)
            with zipfile.ZipFile(output) as archive:
                document_xml = archive.read("word/document.xml")
                document = document_xml.decode("utf-8")
                settings = archive.read("word/settings.xml").decode("utf-8")

        paragraphs, ns = self._xml_paragraphs(document_xml)
        toc_titles = [
            paragraph for paragraph in paragraphs if "自定义目录" in "".join(
                node.text or "" for node in paragraph.findall(".//w:t", ns)
            )
        ]
        self.assertEqual(len(toc_titles), 1)
        if len(toc_titles) != 1:
            return
        toc_title = toc_titles[0]
        title_fonts = toc_title.find(".//w:rPr/w:rFonts", ns)
        title_size = toc_title.find(".//w:rPr/w:sz", ns)
        self.assertEqual(title_fonts.get(qn("w:eastAsia")), "目录标题字体")
        self.assertEqual(title_size.get(qn("w:val")), "36")
        self.assertIn('w:instrText xml:space="preserve"> TOC \\o "2-4" \\h \\z \\u </w:instrText>', document)
        self.assertIn(">自定义目录</w:t>", document)
        self.assertNotIn("w:updateFields", settings)

    def test_guangdong_writer_does_not_keep_template_config_after_exception(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "custom-body.md"
            template.write_text(
                "+++\n"
                "[body]\n"
                "size_pt = 17\n"
                "+++\n"
                "# 五、交通安全设施技术状况检测评价情况\n",
                encoding="utf-8",
            )
            blocked_output = root / "blocked-output"
            blocked_output.write_text("not a directory", encoding="utf-8")
            with self.assertRaises(OSError):
                self._write_report(root, template, blocked_output)

        self.assertNotIn("_template_format_config", vars(engine.GuangdongChapterWriter))
        self.assertEqual(
            engine.GuangdongChapterWriter._format_config()["body"]["size_pt"],
            markdown_skeleton.DEFAULT_CONFIG["body"]["size_pt"],
        )

    def test_guangdong_two_level_header_cells_have_center_v_align(self) -> None:
        document = Document()
        detail = [{
            "indicator": "bolt",
            "route": "G1",
            "gtype": "二波",
            "position": "左侧",
            "direction": "上行",
            "segment": "K1+000～K2+000",
            "msplice": 1,
            "mconn": 2,
            "asplice": 1,
            "aconn": 2,
            "remark": "",
        }]
        with minimal_docx.format_context(markdown_skeleton.DEFAULT_CONFIG):
            engine.GuangdongChapterWriter._comparison_table(
                document,
                detail,
                {"marking": 7, "height": 5, "bolt": 5},
            )

        table = document.tables[1]
        xml_rows = table._tbl.findall("./w:tr", table._tbl.nsmap)
        header_cells = [cell for row in xml_rows[:2] for cell in row.findall("./w:tc", table._tbl.nsmap)]
        self.assertEqual([len(row.findall("./w:tc", table._tbl.nsmap)) for row in xml_rows[:2]], [8, 10])
        for cell in header_cells:
            vertical_alignment = cell.find("./w:tcPr/w:vAlign", cell.nsmap)
            self.assertIsNotNone(vertical_alignment)
            self.assertEqual(vertical_alignment.get(qn("w:val")), "center")
        self.assertEqual(
            sum(cell.find("./w:tcPr/w:gridSpan", cell.nsmap) is not None for cell in xml_rows[0].findall("./w:tc", table._tbl.nsmap)),
            2,
        )
        self.assertEqual(
            sum(cell.find("./w:tcPr/w:vMerge", cell.nsmap) is not None for cell in header_cells),
            12,
        )

    def test_guangdong_toc_field_paragraph_has_no_body_first_line_indent(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "toc.md"
            template.write_text(
                "+++\n"
                "[toc]\n"
                "enabled = true\n"
                "title = \"目录标题\"\n"
                "+++\n"
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "<!-- toc -->\n",
                encoding="utf-8",
            )
            output = self._write_report(root, template)
            with zipfile.ZipFile(output) as archive:
                document_xml = archive.read("word/document.xml")

        paragraphs, ns = self._xml_paragraphs(document_xml)
        field_paragraph = next(
            paragraph for paragraph in paragraphs if paragraph.find(".//w:instrText", ns) is not None
        )
        style = field_paragraph.find("./w:pPr/w:pStyle", ns)
        indent = field_paragraph.find("./w:pPr/w:ind", ns)
        self.assertIsNotNone(style)
        self.assertTrue(style.get(qn("w:val")).startswith("TOC"))
        self.assertNotEqual(None if indent is None else indent.get(qn("w:firstLine")), "420")


class SharedRulesTests(unittest.TestCase):
    def test_guardrail_height_and_bolt_rules_are_shared(self) -> None:
        self.assertEqual(engine.guardrail_type("两波护栏"), "二波")
        self.assertEqual(engine.guardrail_type("三波护栏"), "三波")
        self.assertTrue(engine.height_deviation_over_10cm("二波", 701))
        self.assertFalse(engine.height_deviation_over_10cm("三波", 797))
        self.assertAlmostEqual(engine.bolt_missing_ratio(10, 20, 3), 3 / 33)
        self.assertIsNone(engine.bolt_missing_ratio(0, 0, 0))

    def test_progress_message_round_trips_with_optional_item(self) -> None:
        message = engine.format_progress("扫描资料", 2, 5, "a.xlsx")
        self.assertEqual(message, "[progress] 扫描资料 2/5 a.xlsx")
        self.assertEqual(
            engine.parse_progress(message),
            {"stage": "扫描资料", "current": 2, "total": 5, "item": "a.xlsx"},
        )
        self.assertIsNone(engine.parse_progress("普通日志"))

    def test_make_excel_keeps_zero_denominator_bolt_rate_blank(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = engine.Config(root, root / "summary.xlsx", root, root / "template.md", root / "output")
            segment = {
                "county": "测试区",
                "route": "G1",
                "start": 1000.0,
                "end": 2000.0,
                "mileage": 1.0,
            }
            engine.make_excel(
                config,
                [segment],
                bolt_stats=[{
                    "segment": segment,
                    "splice": 0,
                    "connection": 0,
                    "missing": 0,
                    "rate": None,
                }],
            )
            workbook = openpyxl.load_workbook(config.out_xlsx, data_only=True)
            self.assertIsNone(workbook["螺栓缺失统计"]["J2"].value)
            workbook.close()

    def test_docx_bolt_sentence_keeps_zero_denominator_rate_blank(self) -> None:
        segment = {
            "county": "测试区",
            "route": "G1",
            "start": 1000.0,
            "end": 2000.0,
            "mileage": 1.0,
        }
        stat = {
            "segment": segment,
            "splice": 0,
            "connection": 0,
            "missing": 0,
            "rate": None,
            "points": 1,
        }
        document = Document()
        with tempfile.TemporaryDirectory() as temp_dir:
            minimal_docx._section_bolt(document, [segment], [stat], [], {}, temp_dir)
        self.assertIn("缺失率为%。", "\n".join(p.text for p in document.paragraphs))

    def test_docx_bolt_zero_points_with_height_data_uses_neutral_sentence(self) -> None:
        # D6：同段高度有有效数据时，不得写“本段无波形护栏”。
        segment = {
            "county": "测试区",
            "route": "G1",
            "start": 1000.0,
            "end": 2000.0,
            "mileage": 1.0,
        }
        stat = {
            "segment": segment,
            "splice": 0,
            "connection": 0,
            "missing": 0,
            "rate": None,
            "points": 0,
        }
        height_stats = [{
            "segment": segment,
            "types": {
                "二波": {"count": 12, "bins": [0, 0, 12, 0, 0], "pcts": [0, 0, 100, 0, 0], "pass": 100},
                "三波": {"count": 0, "bins": [0, 0, 0, 0, 0], "pcts": [0, 0, 0, 0, 0], "pass": 0},
            },
        }]
        document = Document()
        with tempfile.TemporaryDirectory() as temp_dir:
            minimal_docx._section_bolt(document, [segment], [stat], [], {}, temp_dir, height_stats)
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("本段螺栓明细无有效记录", text)
        self.assertNotIn("本段无波形护栏", text)

    def test_docx_bolt_zero_points_without_height_data_keeps_no_guardrail_sentence(self) -> None:
        # D6：高度同样无数据时，保留“本段无波形护栏”判断。
        segment = {
            "county": "测试区",
            "route": "G1",
            "start": 1000.0,
            "end": 2000.0,
            "mileage": 1.0,
        }
        stat = {
            "segment": segment,
            "splice": 0,
            "connection": 0,
            "missing": 0,
            "rate": None,
            "points": 0,
        }
        height_stats = [{
            "segment": segment,
            "types": {
                "二波": {"count": 0, "bins": [0, 0, 0, 0, 0], "pcts": [0, 0, 0, 0, 0], "pass": 0},
                "三波": {"count": 0, "bins": [0, 0, 0, 0, 0], "pcts": [0, 0, 0, 0, 0], "pass": 0},
            },
        }]
        document = Document()
        with tempfile.TemporaryDirectory() as temp_dir:
            minimal_docx._section_bolt(document, [segment], [stat], [], {}, temp_dir, height_stats)
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("本段无波形护栏", text)


    def test_iter_height_rows_resolves_shared_string_headers(self) -> None:
        parts = {
            "[Content_Types].xml": """<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>""",
            "_rels/.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>""",
            "xl/workbook.xml": """<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>
</workbook>""",
            "xl/_rels/workbook.xml.rels": """<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>""",
            "xl/sharedStrings.xml": """<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">
<si><t>护栏类型</t></si><si><t>梁板中心高度(mm)</t></si><si><t>原始桩号</t></si><si><t>三波护栏</t></si>
</sst>""",
            "xl/worksheets/sheet1.xml": """<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1" t="s"><v>2</v></c></row>
<row r="2"><c r="A2" t="s"><v>3</v></c><c r="B2"><v>697</v></c><c r="C2"><v>2101.167</v></c></row>
</sheetData></worksheet>""",
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "shared-strings.xlsx"
            with zipfile.ZipFile(path, "w") as archive:
                for name, content in parts.items():
                    archive.writestr(name, content)
            rows = list(engine.iter_height_rows(path))
        self.assertEqual(rows[0]["护栏类型"], "三波护栏")
        self.assertEqual(rows[0]["梁板中心高度(mm)"], 697.0)

    def test_segment_lookup_requires_source_route_when_routes_overlap(self) -> None:
        segments = [
            {"route": "G319", "start": 2100000.0, "end": 2200000.0},
            {"route": "G210", "start": 2100000.0, "end": 2200000.0},
        ]
        self.assertEqual(engine._segment_index(segments, 2150000.0, "G210"), 1)
        self.assertIsNone(engine._segment_index(segments, 2150000.0, "G351"))

    def test_tci_rows_use_data_route_column_not_header_label(self) -> None:
        segments = [
            {"route": "G319", "start": 2100000.0, "end": 2200000.0},
            {"route": "G210", "start": 2100000.0, "end": 2200000.0},
        ]
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "tci.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["区域", "路线编号", "原始桩号", "电子修正桩号", "防护设施缺损", "标志缺损", "标线缺损"])
            sheet.append(["", "", "", "", "轻", "", ""])
            sheet.append(["测试区", "G210", "K2150+000", "", 1, 0, 0])
            workbook.save(path)
            records = engine.collect_tci_records(segments, path)

        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["segment"], 1)


class DesktopBridgeTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.project = self.root / "project"
        self.project.mkdir()
        self.output = self.root / "output"
        self.summary = self.root / "summary.xlsx"
        self.manual = self.root / "manual.xlsx"
        self.route = self.root / "route.xlsx"
        for path in (self.summary, self.manual, self.route):
            path.write_bytes(b"test")
        self.detail = self.root / "detail"
        self.disease = self.root / "disease"
        self.tci = self.root / "tci"
        self.detail.mkdir()
        self.disease.mkdir()
        self.tci.mkdir()
        self.cq_template = self.root / "重庆项目报告模板.md"
        self.gd_template = self.root / "广东项目第五章模板.md"
        self.cq_template.write_text("# 重庆模板\n", encoding="utf-8")
        self.gd_template.write_text("# 广东模板\n", encoding="utf-8")
        self.bridge = DesktopBridge()

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def templates(self) -> dict[str, Path]:
        return {
            "重庆项目报告模板": self.cq_template,
            "广东项目第五章模板": self.gd_template,
        }

    def cq_payload(self) -> dict:
        return {
            "template": "cq",
            "values": {
                "projectPath": str(self.project),
                "summaryPath": str(self.summary),
                "detailPath": str(self.detail),
                "diseasePath": str(self.disease),
                "tciPath": str(self.tci),
                "outputPath": str(self.output),
            },
        }

    def test_native_window_reference_is_private_from_js_api(self) -> None:
        window = FakeWindow()
        self.bridge.attach_window(window)
        self.assertNotIn("window", vars(self.bridge))
        self.assertIs(self.bridge._window, window)

    def test_validate_chongqing_payload_creates_output(self) -> None:
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            self.bridge._validate_payload(self.cq_payload())
        self.assertTrue(self.output.is_dir())

    def test_start_run_validation_error_releases_lock(self) -> None:
        payload = self.cq_payload()
        payload["values"]["summaryPath"] = ""
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            result = self.bridge.start_run(payload)
        self.assertFalse(result["ok"])
        self.assertIn("分段汇总表", result["error"])
        self.assertFalse(self.bridge._running)

    def test_chongqing_payload_requires_tci_folder(self) -> None:
        payload = self.cq_payload()
        payload["values"]["tciPath"] = ""
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            result = self.bridge.start_run(payload)
        self.assertFalse(result["ok"])
        self.assertIn("TCI数据文件夹", result["error"])
        self.assertFalse(self.bridge._running)

    def test_chongqing_run_wires_tci_path_and_process_tci(self) -> None:
        configs = []
        calls = []

        def fake_run(config, **kwargs):
            configs.append(config)
            calls.append(kwargs)

        with patch.object(self.bridge, "_template_paths", return_value=self.templates()), \
                patch.object(engine, "generate_statistics_and_report", side_effect=fake_run):
            self.bridge._run_chongqing(self.cq_payload()["values"])
        self.assertEqual(str(configs[0].tci_path), str(self.tci))
        self.assertTrue(calls[0]["process_tci"])

    def test_discover_paths_detects_tci_folder(self) -> None:
        tci_dir = self.project / "TCI数据"
        tci_dir.mkdir()
        (tci_dir / "万州区-G348.xlsx").write_bytes(b"x")
        summary, detail, disease, tci = engine.discover_paths(self.project)
        self.assertEqual(tci, tci_dir)

    def test_guangdong_specialized_folders_are_optional(self) -> None:
        payload = {
            "template": "gd",
            "values": {
                "projectPath": str(self.project),
                "markingPath": "",
                "guardrailPath": "",
                "manualPath": str(self.manual),
                "routePath": str(self.route),
                "outputPath": str(self.output),
                "markingThreshold": "7",
                "heightThreshold": "5",
                "boltThreshold": "5",
            },
        }
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            self.bridge._validate_payload(payload)
        self.assertTrue(self.output.is_dir())

    def test_missing_specialized_template_fails_guangdong_validation(self) -> None:
        templates = self.templates()
        templates["广东项目第五章模板"] = self.root / "missing-gd.md"
        scalar = {
            "template": "gd",
            "values": {
                "projectPath": str(self.project),
                "markingPath": "",
                "guardrailPath": "",
                "manualPath": str(self.manual),
                "routePath": str(self.route),
                "outputPath": str(self.output),
                "markingThreshold": "7",
                "heightThreshold": "5",
                "boltThreshold": "5",
            },
        }
        with patch.object(self.bridge, "_template_paths", return_value=templates):
            with self.assertRaisesRegex(FileNotFoundError, "templates 文件夹"):
                self.bridge._validate_payload(scalar)

    def test_worker_emits_complete_and_releases_running_state(self) -> None:
        window = FakeWindow()
        self.bridge.attach_window(window)
        self.bridge._running = True
        with patch.object(self.bridge, "_run_chongqing"):
            self.bridge._run_worker(self.cq_payload())
        self.assertFalse(self.bridge._running)
        events = [json.loads(script.split("desktopEvents(", 1)[1][:-1]) for script in window.scripts]
        self.assertEqual(events[0]["status"], "running")
        self.assertEqual(events[-1]["status"], "complete")
        self.assertEqual(events[-1]["progress"], 100)

    def test_progress_mapping(self) -> None:
        self.assertEqual(self.bridge._progress_from_log("正在扫描资料"), (34, 1))
        self.assertEqual(self.bridge._progress_from_log("图表已生成"), (72, 2))
        self.assertEqual(self.bridge._progress_from_log("已保存文件"), (92, 3))

    def test_structured_progress_mapping_is_incremental(self) -> None:
        first = self.bridge._progress_from_log(engine.format_progress("扫描资料", 1, 4))
        last = self.bridge._progress_from_log(engine.format_progress("扫描资料", 4, 4))
        self.assertEqual(first[1], 1)
        self.assertEqual(last[1], 1)
        self.assertLess(first[0], last[0])

    def test_resource_template_path_uses_application_root(self) -> None:
        self.assertEqual(
            engine.resource_template_path("x.md"),
            engine.application_root() / "templates" / "x.md",
        )

    def test_guangdong_run_uses_resource_template_for_bundles(self) -> None:
        values = {
            "projectPath": str(self.project),
            "markingPath": "",
            "guardrailPath": "",
            "manualPath": str(self.manual),
            "routePath": str(self.route),
            "outputPath": str(self.output),
            "markingThreshold": "7",
            "heightThreshold": "5",
            "boltThreshold": "5",
        }
        scanned = {
            "marking": [{"city": "佛山市", "route": "G1"}],
            "height": [],
            "bolt": [],
            "notes": [],
            "issues": [],
        }
        expected_template = self.root / "resource-template.md"
        expected_template.write_text("# 模板", encoding="utf-8")
        route_index = type("RouteIndex", (), {"mapping": {}})()
        with (
            patch.object(engine.RouteCategoryIndex, "from_file", return_value=route_index),
            patch.object(engine.GuangdongInputScanner, "scan", return_value=scanned),
            patch.object(engine.ManualAutoComparator, "read_file", return_value=([], [])),
            patch.object(engine.GuangdongBatchRunner, "build_bundles", return_value={}),
            patch.object(engine.GuangdongBatchRunner, "run_bundles", return_value={}) as run_bundles,
            patch.object(engine, "resource_template_path", return_value=expected_template) as resource_path,
        ):
            self.bridge._run_guangdong(values)

        resource_path.assert_called_once_with("广东项目第五章模板.md")
        self.assertEqual(run_bundles.call_args.args[2], expected_template)


class UpstreamCountyTests(unittest.TestCase):
    def test_read_segments_reads_county_column(self) -> None:
        import openpyxl

        from backend.report_engine import read_segments

        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
            ws.append([1, "两江新区", "G210", "满都拉－防城港", "一级公路", 2157.392, 2159.964, 2.572, 37.476])
            ws.append([1, "两江新区", "G210", "满都拉－防城港", "一级公路", 2159.964, 2184.977, 25.013, 37.476])
            wb.save(tmp_path / "summary.xlsx")
            segs = read_segments(tmp_path / "summary.xlsx")
            self.assertEqual(segs[0]["county"], "两江新区")
            self.assertEqual(segs[0]["route"], "G210")
            self.assertEqual(segs[0]["route_name"], "满都拉－防城港")


def build_demo_segments():
    return [
        {
            "grade": "一级",
            "manager": "重庆交通行政执法总队",
            "start": 2264000.0,
            "end": 2265000.0,
            "mileage": 1.0,
        },
        {
            "grade": "一级",
            "manager": "重庆交通行政执法总队",
            "start": 2265000.0,
            "end": 2266000.0,
            "mileage": 1.0,
        },
    ]


def build_demo_height_records():
    records = []
    for index, (start, end) in enumerate([(2264000.0, 2265000.0), (2265000.0, 2266000.0)]):
        for offset in (10, 20, 30):
            records.append({
                "file": "交安设施现场检测明细.xlsx",
                "direction": "上行",
                "station": start + offset * 10,
                "raw_station": start + offset * 10,
                "electronic_station": start + offset * 10,
                "basis": "电子修正桩号",
                "kind": "二波",
                "height": 580.0 + (offset * 2 % 40),
                "segment": index,
            })
        for offset in (40, 50):
            records.append({
                "file": "交安设施现场检测明细.xlsx",
                "direction": "上行",
                "station": start + offset * 10,
                "raw_station": start + offset * 10,
                "electronic_station": start + offset * 10,
                "basis": "电子修正桩号",
                "kind": "三波",
                "height": 677.0 + (offset % 20),
                "segment": index,
            })
    return records


def build_demo_bolt_records():
    return [
        {
            "file": "交安设施现场检测明细.xlsx",
            "direction": "下行",
            "station": 2264010.0,
            "raw_station": 2264010.0,
            "electronic_station": 2264010.0,
            "basis": "原始桩号",
            "segment": 0,
            "splice": 80,
            "splice_missing": 3,
            "connection": 120,
            "connection_missing": 1,
        }
    ]


def build_demo_stats(segments, records):
    stats = []
    for index in range(len(segments)):
        kinds = {"二波": [], "三波": []}
        bins = {"二波": [0] * 5, "三波": [0] * 5}
        for record in records:
            if record["segment"] != index:
                continue
            kinds[record["kind"]].append(record["height"])
            bin_index = report_engine_bin(record)
            bins[record["kind"]][bin_index] += 1
        types = {}
        for kind in ("二波", "三波"):
            heights = kinds[kind]
            pcts = [round(v * 100 / len(heights), 2) if heights else 0 for v in bins[kind]]
            types[kind] = {
                "count": len(heights),
                "bins": bins[kind],
                "pcts": pcts,
                "pass": pcts[2] if heights else 0,
            }
        stats.append({"segment": segments[index], "types": types})
    return stats


def report_engine_bin(record):
    kind = record["kind"]
    if kind == "二波":
        limits = (550, 580, 620, 650)
    else:
        limits = (647, 677, 717, 747)
    height = record["height"]
    if height < limits[0]:
        return 0
    if height < limits[1]:
        return 1
    if height <= limits[2]:
        return 2
    if height <= limits[3]:
        return 3
    return 4


class UpdaterTests(unittest.TestCase):
    def test_updater_core_validation(self) -> None:
        from updater import is_newer, parse_version, select_installer_asset, validate_download

        self.assertEqual(parse_version("v1.2.3"), (1, 2, 3))
        self.assertTrue(is_newer((1, 2, 4), (1, 2, 3)))
        self.assertFalse(is_newer((1, 2, 3), (1, 2, 3)))
        asset = select_installer_asset({
            "assets": [{
                "name": "report-generator-Setup.exe",
                "browser_download_url": "https://github.com/a/b/releases/download/v1/报告生成工具-Setup.exe",
                "size": 2,
            }],
        })
        self.assertEqual(asset["name"], "report-generator-Setup.exe")
        with self.assertRaises(ValueError):
            validate_download(b"not-an-exe", 11)


    def test_updater_rejects_untrusted_assets_and_checks_release(self) -> None:
        from updater import check_for_update, select_installer_asset

        release = {
            "tag_name": "v0.2.0",
            "assets": [{
                "name": "report-generator-Setup.exe",
                "browser_download_url": "https://objects.githubusercontent.com/a/b.exe",
                "size": 2,
            }],
        }
        result = check_for_update("0.1.0", lambda: release)
        self.assertTrue(result["update_available"])
        self.assertEqual(result["latest_version"], "0.2.0")

        malicious = dict(release)
        malicious["assets"] = [{
            "name": "report-generator-Setup.exe",
            "browser_download_url": "http://example.com/update.exe",
            "size": 2,
        }]
        with self.assertRaises(ValueError):
            select_installer_asset(malicious)

    def test_updater_launcher_does_not_use_shell(self) -> None:
        from updater import launch_installer

        with patch("updater.subprocess.Popen") as popen:
            launch_installer(Path("C:/Temp/update.exe"))
        popen.assert_called_once_with([str(Path("C:/Temp/update.exe"))], close_fds=True)


class MinimalDocxTests(unittest.TestCase):
    def test_generates_report_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir)
            skeleton = Path(temp_dir) / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "## 3.1 G210线整体情况\n\n"
                "## 5.1 G210线整体情况\n\n"
                "## 6 结论与建议\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=Path(temp_dir),
                summary_xlsx=Path(temp_dir) / "summary.xlsx",
                detail_dir=Path(temp_dir),
                template_docx=skeleton,
                output_dir=output,
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            messages = []
            config.out_docx.unlink(missing_ok=True)
            result = minimal_docx.run(
                config,
                segments,
                height_stats,
                records,
                bolt_stats,
                bolt_records,
                None,
                log=messages.append,
                skeleton_md=skeleton,
            )
            self.assertTrue(result.is_file())
            self.assertGreater(result.stat().st_size, 12000)
            self.assertTrue(any("Markdown" in message or "程序化" in message for message in messages))

    def test_engine_routes_to_minimal_without_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            skeleton = root / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "## 3.1 G210线整体情况\n\n"
                "## 5.1 G210线整体情况\n\n"
                "## 6 结论与建议\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=skeleton,
                output_dir=output,
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            messages = []
            engine.make_docx(
                config,
                segments,
                height_stats=height_stats,
                height_records=records,
                bolt_stats=bolt_stats,
                bolt_records=bolt_records,
                disease_image_index=None,
                log=messages.append,
                require_template=False,
            )
            self.assertTrue(config.out_docx.is_file())
            self.assertGreater(config.out_docx.stat().st_size, 12000)
            self.assertTrue(any("Markdown" in message or "程序化" in message for message in messages))

    def test_generates_report_from_markdown_skeleton(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            skeleton = root / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "<!-- inject:overview -->\n\n"
                "本报告依据委托单位提供的数据生成。\n\n"
                "## 1.2 检测依据\n\n"
                "依据《公路技术状况评定标准》开展检测。\n\n"
                "# 3 沿线设施技术状况评价\n\n"
                "<!-- inject:tci -->\n\n"
                "# 4 波形梁护栏横梁中心高度检测结果\n\n"
                "<!-- inject:height -->\n\n"
                "### 3.2 G210线K2264+000~K2265+000段（样例）\n\n"
                "# 5 波形梁护栏螺栓缺失\n\n"
                "<!-- inject:bolt -->\n\n"
                "# 6 结论与建议\n\n"
                "<!-- inject:conclusion -->\n\n"
                "## 6 建议\n\n"
                "标志牌安装情况检查合格。\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=skeleton,
                output_dir=root / "output",
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            result = minimal_docx.run(
                config,
                segments,
                height_stats,
                records,
                bolt_stats,
                bolt_records,
                None,
                skeleton_md=skeleton,
            )
            self.assertTrue(result.is_file())
            document = Document(result)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            self.assertIn("自动化检测报告", text)
            self.assertIn("本报告依据委托单位提供的数据生成。", text)
            self.assertIn("波形梁护栏横梁中心高度检测结果", text)
            self.assertIn("结论与建议", text)
            self.assertNotIn("（样例）", text)

    def test_conclusion_renders_once_despite_static_template_leftovers(self) -> None:
        # D2：即使模板残留静态 5.1结论/5.2建议，结论/建议也只渲染一轮。
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            skeleton = root / "重庆项目报告模板.md"
            skeleton.write_text(
                "# 2026年普通公路国省道交通安全设施自动化检测报告\n\n"
                "## 1.1 项目概况\n\n"
                "# 5 结论与建议\n\n"
                "<!-- inject:conclusion -->\n\n"
                "## 5.1 结论\n\n"
                "结论章节由程序根据检测统计自动生成。\n\n"
                "## 5.2 建议\n\n"
                "建议章节由程序根据检测统计自动生成。\n",
                encoding="utf-8",
            )
            config = engine.Config(
                project_dir=root,
                summary_xlsx=root / "summary.xlsx",
                detail_dir=root,
                template_docx=skeleton,
                output_dir=root / "output",
                disease_dir=None,
            )
            segments = build_demo_segments()
            records = build_demo_height_records()
            bolt_records = build_demo_bolt_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, bolt_records)
            result = minimal_docx.run(
                config,
                segments,
                height_stats,
                records,
                bolt_stats,
                bolt_records,
                None,
                skeleton_md=skeleton,
            )
            self.assertTrue(result.is_file())
            document = Document(result)
            headings = [p.text for p in document.paragraphs if p.style.name.startswith("Heading")]
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)
            self.assertEqual(headings.count("总结"), 1, headings)
            self.assertEqual(headings.count("建议"), 1, headings)
            self.assertFalse(any("5.1" in h or "5.2" in h for h in headings), headings)
            self.assertNotIn("由程序根据检测统计自动生成", text)
            # D6 联动：第二段螺栓无记录但高度有数据，应为中性表述。
            self.assertIn("G210线共检出拼接螺栓80颗", text)
            self.assertNotIn("本段无波形护栏", text)

    def test_overview_table_has_county_column(self) -> None:
        segs = [{"county": "两江新区", "route": "G210", "route_name": "满都拉－防城港", "grade": "一级公路", "start": 2157392, "end": 2159964, "mileage": 2.572, "total_mileage": 37.476, "manager": ""}]
        doc = Document()
        minimal_docx._section_overview(doc, None, segs)
        table = doc.tables[0]
        headers = [c.text for c in table.rows[0].cells]
        self.assertEqual(headers[1], "区县")
        self.assertEqual(headers[3], "路线名")
        self.assertEqual(len(headers), 9)
        self.assertEqual(headers[0], "序号")
        row = [c.text for c in table.rows[1].cells]
        self.assertEqual(row[1], "两江新区")
        self.assertEqual(row[3], "满都拉－防城港")

    def test_height_grouped_by_county(self) -> None:
        segments = [
            {"county": "两江新区", "route": "G210", "route_name": "", "grade": "一级", "manager": "", "start": 2264000.0, "end": 2265000.0, "mileage": 1.0, "total_mileage": 1.0},
            {"county": "北碚区", "route": "G210", "route_name": "", "grade": "一级", "manager": "", "start": 2265000.0, "end": 2266000.0, "mileage": 1.0, "total_mileage": 1.0},
        ]
        records = build_demo_height_records()
        height_stats = build_demo_stats(segments, records)
        bolt_records = [
            {"file": "交安设施现场检测明细.xlsx", "direction": "下行", "station": 2264010.0, "raw_station": 2264010.0, "electronic_station": 2264010.0, "basis": "原始桩号", "segment": 0, "splice": 80, "splice_missing": 3, "connection": 120, "connection_missing": 1},
            {"file": "交安设施现场检测明细.xlsx", "direction": "下行", "station": 2265010.0, "raw_station": 2265010.0, "electronic_station": 2265010.0, "basis": "原始桩号", "segment": 1, "splice": 80, "splice_missing": 3, "connection": 120, "connection_missing": 1},
        ]
        bolt_stats = engine.make_bolt_stats(segments, bolt_records)
        with tempfile.TemporaryDirectory() as tmp:
            import matplotlib.pyplot as plt

            images = {}
            for idx in range(len(segments)):
                for kind in ("二波", "三波"):
                    if height_stats[idx]["types"][kind]["count"] == 0:
                        continue
                    fig, ax = plt.subplots(figsize=(1, 1))
                    ax.plot([0, 1], [0, 1])
                    line_path = Path(tmp) / f"line_{idx}_{kind}.png"
                    fig.savefig(line_path)
                    plt.close(fig)
                    fig2, ax2 = plt.subplots(figsize=(1, 1))
                    ax2.pie([1, 1])
                    pie_path = Path(tmp) / f"pie_{idx}_{kind}.png"
                    fig2.savefig(pie_path)
                    plt.close(fig2)
                    images[(idx, kind)] = {"line": line_path, "pie": pie_path}
            doc = Document()
            minimal_docx._section_height(doc, segments, height_stats, records, images, tmp)
            headings = [p.text for p in doc.paragraphs if p.style.name.startswith("Heading")]
            self.assertTrue(any("两江新区整体情况" in h for h in headings), f"missing 两江新区整体情况 in {headings}")
            self.assertTrue(any("北碚区整体情况" in h for h in headings), f"missing 北碚区整体情况 in {headings}")
            doc2 = Document()
            minimal_docx._section_bolt(doc2, segments, bolt_stats, bolt_records, None, tmp)
            headings2 = [p.text for p in doc2.paragraphs if p.style.name.startswith("Heading")]
            self.assertTrue(any("两江新区整体情况" in h for h in headings2), f"bolt missing 两江新区整体情况 in {headings2}")
            self.assertTrue(any("北碚区整体情况" in h for h in headings2), f"bolt missing 北碚区整体情况 in {headings2}")
            tables = doc.tables
            has_county_col = any(any("区县" in c.text for c in t.rows[0].cells) for t in tables) if tables else False
            tables2 = doc2.tables
            has_county_col2 = any(any("区县" in c.text for c in t.rows[0].cells) for t in tables2) if tables2 else False
            self.assertTrue(has_county_col, f"height tables missing 区县 column, headers: {[[c.text for c in t.rows[0].cells] for t in tables]}")
            self.assertTrue(has_county_col2, f"bolt tables missing 区县 column, headers: {[[c.text for c in t.rows[0].cells] for t in tables2]}")


class GuangdongBusinessRegressionTests(unittest.TestCase):
    def test_scanner_preserves_bridge_guardrail_note(self) -> None:
        scanner = engine.GuangdongInputScanner(".")
        note_row = scanner._convert(
            "height",
            {
                "地市": "佛山市",
                "路线编号": "G1",
                "方向": "上行",
                "电子修正桩号": "K1+000",
                "检测区段": "K1+000～K2+000",
                "护栏类型": "二波",
                "梁板中心高度(mm)": "",
                "备注标记": "桥梁地段",
            },
            Path("guardrail.xlsx"),
            "Sheet1",
        )
        self.assertIsNotNone(note_row)
        self.assertEqual(note_row["guardrail_note"], "桥梁地段")

    def test_writer_handles_single_wave_and_bridge_only_segments(self) -> None:
        def height(segment: str, kind: str, value: float) -> dict:
            return {
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": segment,
                "guardrail_type": kind,
                "height": value,
            }

        def note(segment: str) -> dict:
            return {
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": segment,
                "guardrail_note": "桥梁地段",
            }

        bundle = {
            "marking": [],
            "height": [
                height("K1+000～K2+000", "二波", 600),
                height("K2+000～K3+000", "三波", 700),
            ],
            "bolt": [],
            "notes": [note("K3+000～K4+000")],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n",
                encoding="utf-8",
            )
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)

        self.assertIn("其中二波护栏", text)
        self.assertIn("其中三波护栏", text)
        self.assertEqual(text.count("当前区段为桥梁路段，无有效检测点位。"), 2)
        self.assertNotIn("区段内无护栏", text)
        self.assertNotIn("共检测0个有效点，其中。", text)

    def test_writer_deduplicates_height_and_note_segments_with_suffix(self) -> None:
        bundle = {
            "marking": [],
            "height": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "guardrail_type": "二波",
                "height": 600,
            }],
            "bolt": [],
            "notes": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000段",
                "guardrail_note": "桥梁地段",
            }],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n",
                encoding="utf-8",
            )
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            headings = []
            height_text = []
            in_height_section = False
            for paragraph in document.paragraphs:
                if paragraph.text == "（2）护栏中心高度":
                    in_height_section = True
                elif paragraph.text == "（3）螺栓安装情况":
                    in_height_section = False
                elif in_height_section:
                    height_text.append(paragraph.text)
                    if (
                        paragraph.style.name == "Heading 5"
                        and "K1+000～K2+000段" in paragraph.text
                    ):
                        headings.append(paragraph.text)

        self.assertEqual(headings, ["a. K1+000～K2+000段"])
        self.assertNotIn("当前区段为桥梁路段，无有效检测点位。", "\n".join(height_text))

    def test_writer_accepts_markdown_template_with_city_placeholder(self) -> None:
        bundle = {
            "marking": [],
            "height": [{
                "category": "高速公路",
                "route": "G1",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "guardrail_type": "二波",
                "height": 600,
            }],
            "bolt": [],
            "notes": [],
            "comparison_detail": [],
            "weak_segments": [],
        }
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(
                "# 五、交通安全设施技术状况检测评价情况\n\n"
                "本章为{{地市}}交通安全设施技术状况检测评价内容。\n\n"
                "## （一）高速公路交安设施技术状况\n\n"
                "### 1.沿线设施技术状况TCI\n",
                encoding="utf-8",
            )
            output = engine.GuangdongChapterWriter.write(
                "佛山市",
                bundle,
                root / "output",
                template,
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)

        self.assertIn("本章为佛山市交通安全设施技术状况检测评价内容。", text)
        self.assertIn("五、交通安全设施技术状况检测评价情况", text)
        self.assertIn("（2）护栏中心高度", text)
        self.assertIn("其中二波护栏", text)
        self.assertNotIn("{{地市}}", text)

    def test_comparison_table_emits_expected_two_level_bolt_header_xml(self) -> None:
        document = Document()
        engine.GuangdongChapterWriter._comparison_table(
            document,
            [{
                "indicator": "bolt",
                "route": "G1",
                "gtype": "二波",
                "position": "左侧",
                "direction": "上行",
                "segment": "K1+000～K2+000",
                "msplice": 1,
                "mconn": 2,
                "asplice": 1,
                "aconn": 2,
                "remark": "",
            }],
            {"marking": 7, "height": 5, "bolt": 5},
        )
        table = document.tables[1]
        xml_rows = table._tbl.findall("./w:tr", table._tbl.nsmap)
        cell_counts = [len(row.findall("./w:tc", table._tbl.nsmap)) for row in xml_rows]
        top_cells = xml_rows[0].findall("./w:tc", table._tbl.nsmap)
        spans = []
        for cell in top_cells:
            span = cell.find("./w:tcPr/w:gridSpan", cell.nsmap)
            spans.append(int(span.get(qn("w:val"))) if span is not None else 1)
        grid_columns = table._tbl.findall("./w:tblGrid/w:gridCol", table._tbl.nsmap)

        self.assertEqual(cell_counts, [8, 10, 10])
        self.assertEqual(len(grid_columns), 10)
        self.assertEqual(spans, [1, 1, 1, 1, 1, 2, 2, 1])

    def test_comparison_analysis_uses_average_without_min_max_range(self) -> None:
        document = Document()
        detail = [
            {
                "indicator": "height",
                "absolute_difference": 1.0,
                "relative_deviation": None,
                "within_threshold": True,
            },
            {
                "indicator": "height",
                "absolute_difference": 3.0,
                "relative_deviation": None,
                "within_threshold": False,
            },
        ]
        engine.GuangdongChapterWriter._comparison_table(
            document,
            detail,
            {"marking": 7, "height": 5, "bolt": 5},
        )
        text = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("平均偏差", text)
        self.assertIn("一致性占比", text)
        self.assertNotIn("1.00～3.00", text)


REAL_E2E_ENV = "REPORT_E2E_REAL"
_REAL_E2E_PATHS = {
    "route_xlsx": ("REPORT_E2E_ROUTE_XLSX", "file"),
    "manual_xlsx": ("REPORT_E2E_MANUAL_XLSX", "file"),
    "template": ("REPORT_E2E_TEMPLATE_DOCX", "file"),
    "foshan_marking": ("REPORT_E2E_FOSHAN_MARKING_DIR", "dir"),
    "foshan_guardrail": ("REPORT_E2E_FOSHAN_GUARDRAIL_DIR", "dir"),
    "zhuhai_marking": ("REPORT_E2E_ZHUHAI_MARKING_DIR", "dir"),
    "zhuhai_guardrail": ("REPORT_E2E_ZHUHAI_GUARDRAIL_DIR", "dir"),
}


def _required_real_path(key: str) -> Path:
    env_name, kind = _REAL_E2E_PATHS[key]
    value = os.environ.get(env_name, "").strip()
    if not value:
        raise AssertionError(f"真实数据测试必须设置环境变量：{env_name}")
    path = Path(value)
    valid = path.is_file() if kind == "file" else path.is_dir()
    if not valid:
        raise AssertionError(f"{env_name} 路径不存在或类型错误：{path}")
    return path


def _real_e2e_inputs() -> dict[str, Path]:
    return {key: _required_real_path(key) for key in _REAL_E2E_PATHS}


class RealDataConfigurationTests(unittest.TestCase):
    def test_real_e2e_requires_input_environment_variables(self) -> None:
        env = {REAL_E2E_ENV: "1"}
        env.update({env_name: "" for env_name, _ in _REAL_E2E_PATHS.values()})
        with patch.dict(os.environ, env, clear=False):
            with self.assertRaisesRegex(AssertionError, "REPORT_E2E_ROUTE_XLSX"):
                _real_e2e_inputs()


def _scan_city(city: dict, route_index) -> dict:
    """真实数据扫描：与 run_guangdong_project 中标线/护栏分别扫描的同款逻辑。"""
    scanned = {"height": [], "bolt": [], "marking": [], "notes": [], "issues": []}
    for key, kinds in (("marking", ("marking",)), ("guardrail", ("height", "bolt", "notes"))):
        root = city.get(key)
        if not root or not root.is_dir():
            continue
        partial = engine.GuangdongInputScanner(root, route_index).scan()
        for kind in kinds:
            scanned[kind].extend(partial.get(kind, []))
        scanned["issues"].extend(partial.get("issues", []))
    return scanned


def _run_real_pipeline(
    cities: list[dict],
    artifact_root: Path,
    route_xlsx: Path,
    manual_xlsx: Path,
    template: Path,
) -> dict:
    """使用运行环境提供的真实输入：直接拼装 scan+build_bundles+run_bundles。"""
    artifact_root.mkdir(parents=True, exist_ok=True)
    route_index = engine.RouteCategoryIndex.from_file(route_xlsx)
    manual_records, manual_issues = engine.ManualAutoComparator.read_file(manual_xlsx)
    thresholds = {"marking": 7, "height": 5, "bolt": 5}

    result = {"success": [], "failed": {}, "warnings": []}
    for city in cities:
        # 每个城市独立 scan/build/run，避免跨城市路线记录进入同一个 bundle。
        scanned = _scan_city(city, route_index)
        scanned["issues"] = list(manual_issues) + scanned["issues"]
        bundles = engine.GuangdongBatchRunner.build_bundles(scanned, route_index, manual_records, thresholds)
        partial = engine.GuangdongBatchRunner.run_bundles(
            bundles, artifact_root, template, thresholds
        )
        result["success"].extend(partial["success"])
        result["failed"].update(partial["failed"])
        result["warnings"].extend(partial["warnings"])
    return result


def _collect_docx_text(docx_path: Path) -> str:
    document = Document(str(docx_path))
    return "\n".join(paragraph.text for paragraph in document.paragraphs)


def _collect_segment_order(text: str, marker: str) -> list[str]:
    """按指标汇总表分段提取 a. 区段标题中的桩号范围。"""
    import re

    summary_markers = (
        "标线逆反射区段汇总表",
        "护栏中心高度区段汇总表",
        "护栏螺栓缺失区段汇总表",
    )
    aliases = {
        "标线逆反射区段": "标线逆反射区段汇总表",
        "护栏中心高度": "护栏中心高度区段汇总表",
        "螺栓": "护栏螺栓缺失区段汇总表",
    }
    wanted = aliases.get(marker, marker)
    if wanted not in summary_markers:
        raise ValueError(f"未知指标汇总表 marker：{marker}")

    title = re.compile(r"^\s*[a-z]\.\s+")
    pattern = re.compile(r"K\+?\d+(?:\.\d+)?[~～]K\+?\d+(?:\.\d+)?")
    result: list[str] = []
    active = False
    for line in text.splitlines():
        found = next((item for item in summary_markers if item in line), None)
        if found is not None:
            active = found == wanted
            continue
        if not active or not title.match(line):
            continue
        match = pattern.search(line)
        if match:
            result.append(match.group(0))
    return result


@unittest.skipUnless(os.environ.get(REAL_E2E_ENV) == "1", f"set {REAL_E2E_ENV}=1 to run")
class RealDataE2ETests(unittest.TestCase):
    """真实佛山/珠海端到端回归；产物仅落在 tests/artifacts/real-data/，不进 git。"""

    @classmethod
    def setUpClass(cls) -> None:
        paths = _real_e2e_inputs()
        cls.route_xlsx = paths["route_xlsx"]
        cls.manual_xlsx = paths["manual_xlsx"]
        cls.template = paths["template"]
        cls.foshan = {
            "name": "佛山",
            "marking": paths["foshan_marking"],
            "guardrail": paths["foshan_guardrail"],
        }
        cls.zhuhai = {
            "name": "珠海",
            "marking": paths["zhuhai_marking"],
            "guardrail": paths["zhuhai_guardrail"],
        }

    def _run(self, cities: list[dict], sub: str) -> dict:
        artifact_root = Path(__file__).resolve().parent / "artifacts" / "real-data" / sub
        if artifact_root.exists():
            # 清空旧产物，避免上轮 run 的 docx 干扰本次断言。
            import shutil

            shutil.rmtree(artifact_root, ignore_errors=True)
        return _run_real_pipeline(cities, artifact_root, self.route_xlsx, self.manual_xlsx, self.template)

    def test_foshan_minimal_integration_runs_end_to_end(self) -> None:
        result = self._run([self.foshan], "foshan")
        self.assertIn("佛山市", result["success"])
        foshan_dir = Path(__file__).resolve().parent / "artifacts" / "real-data" / "foshan" / "佛山市"
        docx_path = next(foshan_dir.glob("*第五部分.docx"), None)
        self.assertIsNotNone(docx_path, f"未生成佛山 docx：{foshan_dir}")
        self._assert_docx_complete(docx_path, "佛山")
        text = _collect_docx_text(docx_path)
        self.assertEqual(text.count("区段内无护栏"), 0)
        self.assertEqual(text.count("共检测0个有效点，其中。"), 0)

    def test_foshan_and_zhuhai_full_e2e_artifact_checks(self) -> None:
        result = self._run([self.foshan, self.zhuhai], "foshan-zhuhai")
        self.assertIn("佛山市", result["success"])
        self.assertIn("珠海市", result["success"], msg=f"珠海失败：{result['failed']}")
        self.assertFalse(result["failed"], f"运行失败：{result['failed']}")

        artifact_root = Path(__file__).resolve().parent / "artifacts" / "real-data" / "foshan-zhuhai"
        foshan_docx = next((artifact_root / "佛山市").glob("*第五部分.docx"), None)
        zhuhai_docx = next((artifact_root / "珠海市").glob("*第五部分.docx"), None)
        self.assertIsNotNone(foshan_docx, "佛山 docx 缺失")
        self.assertIsNotNone(zhuhai_docx, "珠海 docx 缺失")
        self._assert_docx_complete(foshan_docx, "佛山")
        self._assert_docx_complete(zhuhai_docx, "珠海")

        # 仅抽取检查中用得到的城市子串，避免佛山段落污染珠海或反之。
        foshan_text = _collect_docx_text(foshan_docx)
        zhuhai_text = _collect_docx_text(zhuhai_docx)

        for label, text in (("佛山", foshan_text), ("珠海", zhuhai_text)):
            self.assertEqual(text.count("区段内无护栏"), 0, f"{label}出现旧句：区段内无护栏")
            self.assertEqual(text.count("共检测0个有效点，其中。"), 0, f"{label}出现残句：共检测0个有效点，其中。")

        # 标线/高度/螺栓三章节共同区段顺序一致：抽取各章节首次出现的区段链，比较相等。
        for label, text in (("佛山", foshan_text), ("珠海", zhuhai_text)):
            for marker in ("标线逆反射区段", "护栏中心高度", "螺栓"):
                segments = _collect_segment_order(text, marker)
                self.assertGreater(len(segments), 1, f"{label}/{marker} 区段数不足，无法比序")
            marking = _collect_segment_order(text, "标线逆反射区段")
            height = _collect_segment_order(text, "护栏中心高度")
            bolt = _collect_segment_order(text, "螺栓")
            common = [s for s in marking if s in set(height) & set(bolt)]
            self.assertGreater(len(common), 1, f"{label} 共同区段数={len(common)}，无法比序")
            height_in_common = [s for s in height if s in set(common)]
            bolt_in_common = [s for s in bolt if s in set(common)]
            self.assertEqual(common, height_in_common, f"{label} 标线/高度共同区段相对顺序不一致")
            self.assertEqual(common, bolt_in_common, f"{label} 标线/螺栓共同区段相对顺序不一致")

        # 抽查珠海 docx 的一个螺栓对比表：物理 w:tc 8/10/10、两个父表头 gridSpan=2。
        bolt_table_index, bolt_table = self._find_bolt_comparison_table(zhuhai_docx)
        with zipfile.ZipFile(zhuhai_docx) as archive:
            xml = archive.read(f"word/document.xml")
        self._assert_two_level_bolt_header(
            xml, bolt_table_index, label=f"珠海 螺栓对比表 #{bolt_table_index}"
        )

    def _find_bolt_comparison_table(self, docx_path: Path):
        document = Document(str(docx_path))
        for index, table in enumerate(document.tables):
            text = "\n".join(cell.text or "" for row in table.rows for cell in row.cells)
            rows = table._tbl.findall("./w:tr", table._tbl.nsmap)
            cell_counts = [len(row.findall("./w:tc", table._tbl.nsmap)) for row in rows]
            if "拼接螺栓" in text and cell_counts[:3] == [8, 10, 10]:
                return index, table
        self.fail(f"未在 {docx_path} 中找到物理列为8/10/10的螺栓对比表")

    def _assert_docx_complete(self, docx_path: Path, label: str) -> None:
        with zipfile.ZipFile(docx_path) as archive:
            self.assertIsNone(archive.testzip(), f"{label} docx ZIP 存在损坏成员")

    def _assert_two_level_bolt_header(self, document_xml: bytes, table_index: int, label: str) -> None:
        from xml.etree import ElementTree as ET

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(document_xml)
        tables = root.findall(f".//{{{ns['w']}}}tbl")
        if not 0 <= table_index < len(tables):
            self.fail(f"{label} XML 表格索引越界：{table_index}/{len(tables)}")
        table = tables[table_index]
        text = "".join(t.text or "" for t in table.iter(f"{{{ns['w']}}}t"))
        self.assertIn("拼接螺栓", text, f"{label} XML 表格内容不符")
        rows = table.findall(f"{{{ns['w']}}}tr")
        cell_counts = [len(r.findall(f"{{{ns['w']}}}tc")) for r in rows]
        self.assertGreaterEqual(len(cell_counts), 3, f"{label} 行数不足：{cell_counts}")
        self.assertEqual(cell_counts[:3], [8, 10, 10], f"{label} 物理列不符：{cell_counts}")
        self.assertTrue(all(count == 10 for count in cell_counts[2:]), f"{label} 数据行物理列不符：{cell_counts}")
        grid = table.find(f"{{{ns['w']}}}tblGrid")
        grid_columns = [] if grid is None else grid.findall(f"{{{ns['w']}}}gridCol")
        self.assertEqual(len(grid_columns), 10, f"{label} 物理网格列不符：{len(grid_columns)}")
        top = rows[0].findall(f"{{{ns['w']}}}tc")
        spans = []
        for cell in top:
            span = cell.find(f"{{{ns['w']}}}tcPr/{{{ns['w']}}}gridSpan")
            spans.append(int(span.get(f"{{{ns['w']}}}val")) if span is not None else 1)
        self.assertEqual(spans.count(2), 2, f"{label} 父表头 gridSpan=2 应有2个：{spans}")
        self.assertEqual(len(spans), 8, f"{label} 父表头数量不符：{spans}")


class ChongqingCountyNamingTests(unittest.TestCase):
    def _config(self, root: Path, county=None) -> object:
        return engine.Config(
            root, root / "summary.xlsx", root, root / "template.md", root / "output",
            county=county,
        )

    def test_default_names_keep_legacy(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            config = self._config(Path(temp_dir))
            self.assertEqual(config.out_docx.name, engine.OUT_DOCX_NAME)
            self.assertEqual(config.out_xlsx.name, engine.OUT_XLSX_NAME)

    def test_county_names_follow_chongqing_pattern(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            config = self._config(Path(temp_dir), county="万州区")
            self.assertEqual(config.out_docx.name, "重庆市万州区交安设施检测报告.docx")
            self.assertEqual(config.out_xlsx.name, "重庆市万州区交安设施检测报告.xlsx")

    def test_county_stem_does_not_duplicate_prefix(self) -> None:
        self.assertEqual(engine.county_report_stem("重庆市万州区"), "重庆市万州区交安设施检测报告")
        self.assertEqual(engine.county_report_stem("城口县"), "重庆市城口县交安设施检测报告")


class CountyRoutingTests(unittest.TestCase):
    def _segments(self) -> list:
        return [
            {"county": "万州区", "route": "G210", "start": 2264000.0, "end": 2265000.0},
            {"county": "渝北区", "route": "G210", "start": 2265000.0, "end": 2266000.0},
        ]

    def test_extract_full_short_and_multi(self) -> None:
        known = ["万州区", "渝北区", "城口县"]
        self.assertEqual(
            engine.extract_counties_from_filenames(["重庆市-万州区-交安设施现场检测-明细.xlsx"], known),
            ["万州区"],
        )
        self.assertEqual(engine.extract_counties_from_filenames(["渝北-G210-明细.xlsx"], known), ["渝北区"])
        self.assertEqual(
            engine.extract_counties_from_filenames(["万州区-明细.xlsx", "城口-病害清单.xlsx"], known),
            ["万州区", "城口县"],
        )
        self.assertEqual(engine.extract_counties_from_filenames(["G210-明细.xlsx"], known), [])

    def test_resolve_without_files_or_dimension_returns_all_or_empty(self) -> None:
        self.assertEqual(
            engine.resolve_report_counties(self._segments(), []),
            ["万州区", "渝北区"],
        )
        self.assertEqual(engine.resolve_report_counties([{"route": "G210"}], ["万州区-明细.xlsx"]), [])

    def test_resolve_without_match_raises_with_candidates(self) -> None:
        with self.assertRaises(ValueError) as raised:
            engine.resolve_report_counties(self._segments(), ["G210-明细.xlsx"])
        self.assertIn("手动选择", str(raised.exception))
        self.assertIn("万州区", str(raised.exception))

    def test_resolve_override_short_ok_and_missing_raises(self) -> None:
        self.assertEqual(engine.resolve_report_counties(self._segments(), [], override="渝北"), ["渝北区"])
        with self.assertRaises(ValueError) as raised:
            engine.resolve_report_counties(self._segments(), [], override="江北区")
        self.assertIn("无对应路线分段", str(raised.exception))


class ChongqingHeadingNumberingTests(unittest.TestCase):
    def _document_xml(self, template_text: str) -> tuple:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(template_text, encoding="utf-8")
            output = root / "output"
            output.mkdir()
            config = engine.Config(root, root / "summary.xlsx", root, template, output)
            minimal_docx.make_report(
                config, [], None, [], None, [], None, [], {}, root,
                skeleton_md=template,
            )
            with zipfile.ZipFile(output / engine.OUT_DOCX_NAME) as archive:
                return (
                    archive.read("word/document.xml").decode("utf-8"),
                    archive.read("word/numbering.xml").decode("utf-8"),
                )

    def _paragraphs_by_style(self, document_xml: str) -> dict:
        from xml.etree import ElementTree as ET

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(document_xml)
        result: dict = {}
        for paragraph in root.findall(".//w:body/w:p", ns):
            style = paragraph.find("w:pPr/w:pStyle", ns)
            if style is None:
                continue
            name = style.get(f"{{{ns['w']}}}val")
            text = "".join(t.text or "" for t in paragraph.findall("w:r/w:t", ns))
            numbered = paragraph.find("w:pPr/w:numPr", ns) is not None
            result.setdefault(name, []).append((text, numbered))
        return result

    def test_multilevel_numbering_and_stripped_text(self) -> None:
        document, numbering = self._document_xml(
            "# 1 概况\n\n## 1.1 项目概况\n\n### 2.3.1 评价\n\n#### 四级标题\n\n##### 五级标题\n"
        )
        self.assertIn('w:val="%1."', numbering)
        self.assertIn('w:val="%1.%2"', numbering)
        self.assertIn('w:val="%1.%2.%3"', numbering)
        by_style = self._paragraphs_by_style(document)
        self.assertEqual(by_style["Heading1"][0][0], "概况")
        self.assertTrue(all(numbered for _, numbered in by_style["Heading1"]))
        self.assertTrue(all(numbered for _, numbered in by_style["Heading2"]))
        self.assertTrue(all(numbered for _, numbered in by_style["Heading3"]))
        self.assertEqual(by_style["Heading3"][0][0], "评价")
        self.assertTrue(all(not numbered for _, numbered in by_style.get("Heading4", [])))
        self.assertTrue(all(not numbered for _, numbered in by_style.get("Heading5", [])))

    def test_year_prefix_not_stripped(self) -> None:
        document, _ = self._document_xml("# 2026年普通公路检测报告\n")
        by_style = self._paragraphs_by_style(document)
        self.assertEqual(by_style["Heading1"][0][0], "2026年普通公路检测报告")

    def test_numbering_symbols_are_black_on_all_levels(self) -> None:
        # T9：抽象编号全部 lvl 的 rPr 直接黑；标题汉字 run 已黑，不动（字体测试另覆）。
        _, numbering = self._document_xml(
            "# 1 概况\n\n## 1.1 项目概况\n\n### 2.3.1 评价\n"
        )
        from xml.etree import ElementTree as ET

        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        root = ET.fromstring(numbering)
        # 默认模板自带多个抽象编号：只校验与 Heading 样式绑定的那一个（我们创建的）。
        ours = [
            abstract for abstract in root.findall(".//w:abstractNum", ns)
            if any(
                (style.get(qn("w:val")) or "").startswith("Heading")
                for style in abstract.findall(".//w:pStyle", ns)
            )
        ]
        self.assertEqual(len(ours), 1)
        lvls = ours[0].findall("w:lvl", ns)
        self.assertEqual(len(lvls), 3)
        for lvl in lvls:
            with self.subTest(ilvl=lvl.get(qn("w:ilvl"))):
                color = lvl.find("w:rPr/w:color", ns)
                self.assertIsNotNone(color)
                self.assertEqual((color.get(qn("w:val")) or "").lower(), "000000")


class ChongqingHeaderFooterTests(unittest.TestCase):
    """T9：重庆页眉页脚 + 封面版式 + E-Mail（只重庆链路，广东 writer 不经过页眉页脚函数）。"""

    _COVER_TEMPLATE = (
        "# @cover 主标题行一|主标题行二|重庆市|报告编号：BG-2026-T9"
        "|项目名称：测试项目|委托单位：测试单位|测试公司|二〇二六年七月\n"
        "\n<!-- toc -->\n\n# 1 概况\n"
    )

    @staticmethod
    def _build(template_text: str) -> tuple:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(template_text, encoding="utf-8")
            output = root / "output"
            output.mkdir()
            config = engine.Config(root, root / "summary.xlsx", root, template, output)
            minimal_docx.make_report(
                config, [], None, [], None, [], None, [], {}, root,
                skeleton_md=template,
            )
            out = output / engine.OUT_DOCX_NAME
            with zipfile.ZipFile(out) as archive:
                parts = {
                    name: archive.read(name).decode("utf-8")
                    for name in archive.namelist()
                    if name.startswith("word/header") or name.startswith("word/footer")
                    or name == "word/document.xml"
                }
            text = "\n".join(paragraph.text for paragraph in Document(out).paragraphs)
            return parts, text

    def test_header_has_company_report_no_and_underline(self) -> None:
        parts, _ = self._build(self._COVER_TEMPLATE)
        headers = {name: xml for name, xml in parts.items() if "/header" in name}
        self.assertTrue(headers, "重庆报告应生成页眉部件")
        body_header = max(headers)  # 正文节页眉（序号最大的页眉部件），首页无页眉引用
        self.assertIn("四川京炜交通工程技术有限公司", headers[body_header])
        self.assertIn("BG-2026-T9", headers[body_header])
        self.assertNotIn("报告编号：", headers[body_header])
        self.assertIn("w:bottom", headers[body_header])
        footers = {name: xml for name, xml in parts.items() if "/footer" in name}
        self.assertTrue(footers, "重庆报告应生成页脚部件")
        body_footer = max(footers)
        self.assertIn("PAGE", footers[body_footer])
        self.assertIn("NUMPAGES", footers[body_footer])
        self.assertIn('w:val="center"', footers[body_footer])

    def test_first_page_clean_and_body_page_number_restarts(self) -> None:
        parts, _ = self._build(self._COVER_TEMPLATE)
        document = parts["word/document.xml"]
        self.assertIn("w:titlePg", document)
        self.assertIn("w:pgNumType", document)
        self.assertIn('w:start="1"', document)

    def test_body_section_has_no_title_pg(self) -> None:
        # T9：正文节不可带 titlePg，否则正文第一页（概况页）无页眉页脚。
        parts, _ = self._build(self._COVER_TEMPLATE)
        root = ET.fromstring(parts["word/document.xml"])
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        sectprs = root.findall(".//w:sectPr", ns)
        self.assertGreaterEqual(len(sectprs), 2, "应有封面节+正文节")
        for sectpr in sectprs[1:]:
            self.assertIsNone(
                sectpr.find("w:titlePg", ns), "正文节不应含 titlePg（首页须显示页眉页脚）"
            )

    def test_cover_has_top_rule_and_underlined_fill_lines(self) -> None:
        parts, text = self._build(self._COVER_TEMPLATE)
        document = parts["word/document.xml"]
        self.assertIn("w:pBdr", document)  # 顶部横线
        self.assertIn('w:u w:val="single"', document)  # 填空值下划线
        self.assertIn("测试项目", text)
        self.assertIn("测试单位", text)

    def test_notes_email_uses_english_prefix(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            output.mkdir()
            template = engine.resource_template_path("重庆项目报告模板.md")
            config = engine.Config(root, root / "summary.xlsx", root, template, output)
            segment = {
                "county": "万州区", "route": "G210", "route_name": "", "grade": "一级",
                "manager": "", "start": 2264000.0, "end": 2265000.0,
                "mileage": 1.0, "total_mileage": 1.0,
            }
            minimal_docx.make_report(
                config, [segment], None, [], None, [], None, [], {}, root,
                skeleton_md=template,
            )
            text = "\n".join(
                paragraph.text
                for paragraph in Document(output / engine.OUT_DOCX_NAME).paragraphs
            )
        self.assertIn("E-Mail：scjwjt@163.com", text)
        self.assertEqual(text.count("Mail："), text.count("E-Mail："))


class ChongqingTableHeaderTests(unittest.TestCase):
    def test_template_fonts_bold_and_gray_shading(self) -> None:
        template = markdown_skeleton.read_template(engine.resource_template_path("重庆项目报告模板.md"))
        self.assertEqual(template.config["body"]["east_asia"], "宋体")
        self.assertEqual(template.config["table"]["east_asia"], "宋体")
        self.assertTrue(template.config["table"]["header_bold"])
        self.assertEqual(template.config["table"]["header_shading"], "D9D9D9")
        for level in ("1", "2", "3"):
            self.assertEqual(template.config["heading"][level]["east_asia"], "黑体")
            self.assertTrue(template.config["heading"][level]["bold"])
        for level in ("4", "5"):
            self.assertEqual(template.config["heading"][level]["east_asia"], "黑体")
            self.assertFalse(template.config["heading"][level]["bold"])

    def test_docx_header_gray_bold(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            output = root / "output"
            output.mkdir()
            template = engine.resource_template_path("重庆项目报告模板.md")
            config = engine.Config(root, root / "summary.xlsx", root, template, output)
            segment = {
                "county": "万州区", "route": "G210", "route_name": "", "grade": "一级",
                "start": 2264000.0, "end": 2265000.0, "mileage": 1.0, "total_mileage": 1.0,
            }
            minimal_docx.make_report(
                config, [segment], None, [], None, [], None, [], {}, root,
                skeleton_md=template,
            )
            with zipfile.ZipFile(output / engine.OUT_DOCX_NAME) as archive:
                document = archive.read("word/document.xml").decode("utf-8")
        self.assertIn('w:fill="D9D9D9"', document)
        self.assertIn("w:eastAsia=\"宋体\"", document)

    def test_excel_header_gray_bold(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = engine.Config(root, root / "summary.xlsx", root, root / "template.md", root / "output")
            segment = {
                "county": "测试区", "route": "G1", "start": 1000.0, "end": 2000.0, "mileage": 1.0,
            }
            engine.make_excel(
                config,
                [segment],
                bolt_stats=[{
                    "segment": segment, "splice": 0, "connection": 0,
                    "missing": 0, "rate": None,
                }],
            )
            workbook = openpyxl.load_workbook(config.out_xlsx, data_only=True)
            header = workbook["螺栓缺失统计"]["A1"]
            self.assertTrue(header.font.bold)
            self.assertTrue(str(header.fill.start_color.rgb).upper().endswith("D9D9D9"))
            self.assertTrue(str(header.font.color.rgb).upper().endswith("000000"))
            workbook.close()


class ChongqingCountyEndToEndTests(unittest.TestCase):
    def _write_summary(self, path: Path) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "各区县项目概况"
        sheet.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
        sheet.append([1, "万州区", "G210", "", "一级", 2264.0, 2265.0, 1.0, 1.0])
        sheet.append([2, "渝北区", "G210", "", "一级", 2265.0, 2266.0, 1.0, 1.0])
        workbook.save(path)

    def _write_detail(self, path: Path, station: str) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["护栏类型", "梁板中心高度(mm)", "原始桩号", "电子修正桩号", "方向", "路线编号", "异常标记"])
        sheet.append(["两波护栏", 600, station, station, "上行", "G210", ""])
        workbook.save(path)

    def _skeleton(self, path: Path) -> None:
        path.write_text("# 报告\n\n## 1.1 项目概况\n", encoding="utf-8")

    def _run(self, root: Path, detail_names: list) -> object:
        summary = root / "summary.xlsx"
        self._write_summary(summary)
        detail = root / "detail"
        detail.mkdir()
        stations = {"wanzhou": "K2264+100", "yubei": "K2265+100"}
        for name in detail_names:
            key = "yubei" if "渝北" in name or "yubei" in name else "wanzhou"
            self._write_detail(detail / name, stations[key])
        skeleton = root / "template.md"
        self._skeleton(skeleton)
        output = root / "output"
        config = engine.Config(root, summary, detail, skeleton, output)
        engine.generate_statistics_and_report(
            config, log=lambda _: None, process_height=True,
            process_bolts=False, process_tci=False, require_template=False,
        )
        return config

    def test_multi_county_files_generate_each_county_report(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            self._run(root, ["重庆市-万州区-交安设施现场检测-明细.xlsx", "渝北-G210-明细.xlsx"])
            for county in ("万州区", "渝北区"):
                stem = engine.county_report_stem(county)
                self.assertTrue((root / "output" / county / f"{stem}.docx").is_file(), county)
                self.assertTrue((root / "output" / county / f"{stem}.xlsx").is_file(), county)
            # R2：多区县时不再生成顶层整体报告
            self.assertFalse((root / "output" / engine.OUT_DOCX_NAME).exists())
            self.assertFalse((root / "output" / engine.OUT_XLSX_NAME).exists())

    def test_single_county_file_filters_segments_and_names_top_files(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            config = self._run(root, ["重庆市-万州区-交安设施现场检测-明细.xlsx"])
            self.assertEqual(config.out_docx.name, "重庆市万州区交安设施检测报告.docx")
            self.assertEqual(config.out_xlsx.name, "重庆市万州区交安设施检测报告.xlsx")
            self.assertTrue(config.out_docx.is_file())
            # R2：单区县时顶层即为该区县报告，不再另建区县子目录副本
            self.assertFalse((root / "output" / "万州区").exists())
            workbook = openpyxl.load_workbook(config.out_xlsx, data_only=True)
            detail = workbook["检测明细"]
            counties = {row[7] for row in detail.iter_rows(min_row=2, values_only=True)}
            workbook.close()
            self.assertEqual(counties, {"万州区"})


class ChongqingFontTimesTests(unittest.TestCase):
    """R1：标题保持黑色，全文英文 Times New Roman（含封面标题）。"""

    def _write_cover_template(self, path: Path) -> None:
        path.write_text(
            "# @cover 主标题行一|主标题行二|重庆市|报告编号：BG-2026-001|项目名称：测试项目|委托单位：测试单位|测试公司|二〇二六年七月\n"
            "\n# 报告章标题\n\n正文段落文字ABC。\n\n<!-- inject:overview -->\n",
            encoding="utf-8",
        )

    def _report_document_xml(self, root: Path, template: Path) -> str:
        output = root / "output"
        output.mkdir()
        config = engine.Config(root, root / "summary.xlsx", root, template, output)
        segment = {
            "county": "万州区", "route": "G210", "route_name": "", "grade": "一级",
            "manager": "", "start": 2264000.0, "end": 2265000.0,
            "mileage": 1.0, "total_mileage": 1.0,
        }
        minimal_docx.make_report(
            config, [segment], None, [], None, [], None, [], {}, root,
            skeleton_md=template,
        )
        with zipfile.ZipFile(output / engine.OUT_DOCX_NAME) as archive:
            return archive.read("word/document.xml").decode("utf-8")

    def test_cover_title_latin_is_times_not_heiti(self) -> None:
        import re

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            self._write_cover_template(template)
            document = self._report_document_xml(root, template)
        self.assertNotIn('w:ascii="黑体"', document)
        self.assertNotIn('w:hAnsi="黑体"', document)
        self.assertIn('w:ascii="Times New Roman"', document)
        self.assertIn('w:val="44"', document)  # 封面主标题 22pt

    def test_all_document_runs_are_black_times(self) -> None:
        import re

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            self._write_cover_template(template)
            document = self._report_document_xml(root, template)
        latins = set(re.findall(r'w:ascii="([^"]+)"', document)) | set(
            re.findall(r'w:hAnsi="([^"]+)"', document)
        )
        self.assertEqual(latins, {"Times New Roman"})
        colors = set(re.findall(r'w:color w:val="([0-9A-Fa-f]+)"', document))
        self.assertEqual(colors, {"000000"})


class ChongqingCountyRouteTests(unittest.TestCase):
    """R3：识别区县→识别道路编号→查汇总表中该道路编号对应分段。"""

    def _segments(self) -> list:
        return [
            {"county": "城口县", "route": "G211", "start": 1000.0, "end": 2000.0},
            {"county": "城口县", "route": "G211", "start": 2000.0, "end": 3000.0},
            {"county": "城口县", "route": "G347", "start": 1000.0, "end": 2000.0},
            {"county": "万州区", "route": "G348", "start": 1000.0, "end": 2000.0},
        ]

    def test_county_route_summary_groups_segments_by_route(self) -> None:
        self.assertEqual(
            engine.county_route_summary(self._segments()),
            {"城口县": {"G211": 2, "G347": 1}, "万州区": {"G348": 1}},
        )

    def test_overlapping_stations_resolve_by_route_within_county(self) -> None:
        segments = self._segments()
        self.assertEqual(engine._segment_index(segments, 1500.0, "G211"), 0)
        self.assertEqual(engine._segment_index(segments, 1500.0, "G347"), 2)
        self.assertEqual(engine._segment_index(segments, 1500.0, "G348"), 3)
        self.assertIsNone(engine._segment_index(segments, 1500.0, "G210"))


class ChongqingIntervalConvergenceTests(unittest.TestCase):
    """R3：区间明细表按区县/区段收敛，不按全量汇总表展开。"""

    def _write_summary(self, path: Path) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "各区县项目概况"
        sheet.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
        sheet.append([1, "万州区", "G210", "", "一级", 2264.0, 2265.0, 1.0, 1.0])
        sheet.append([2, "渝北区", "G210", "", "一级", 2265.0, 2266.0, 1.0, 1.0])
        workbook.save(path)

    def test_add_interval_sheets_converges_to_given_segments(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            summary = root / "summary.xlsx"
            self._write_summary(summary)
            segments = engine.read_segments(summary)
            county_segments = [s for s in segments if s["county"] == "万州区"]
            record = {
                "file": "f.xlsx", "direction": "上行", "station": 2264100.0,
                "raw_station": 2264100.0, "electronic_station": 2264100.0,
                "basis": "电子修正桩号", "kind": "二波", "height": 600.0, "segment": 0,
            }
            height_stats = engine.make_stats(county_segments, [record])
            config = engine.Config(
                root, summary, root, root / "template.md", root / "output",
                county="万州区",
            )
            engine.make_excel(
                config, county_segments,
                height_stats=height_stats, height_records=[record],
            )
            workbook = openpyxl.load_workbook(config.out_xlsx, data_only=True)
            intervals = [name for name in workbook.sheetnames if name.startswith("区间")]
            count = workbook[intervals[0]]["K2"].value if intervals else None
            workbook.close()
        self.assertEqual(len(intervals), 1)
        self.assertIn("2264+000", intervals[0])
        self.assertEqual(count, 1)

    def test_add_charts_places_pies_in_county_mode(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            summary = root / "summary.xlsx"
            self._write_summary(summary)
            segments = engine.read_segments(summary)
            county_segments = [s for s in segments if s["county"] == "万州区"]
            record = {
                "file": "f.xlsx", "direction": "上行", "station": 2264100.0,
                "raw_station": 2264100.0, "electronic_station": 2264100.0,
                "basis": "电子修正桩号", "kind": "二波", "height": 600.0, "segment": 0,
            }
            height_stats = engine.make_stats(county_segments, [record])
            config = engine.Config(
                root, summary, root, root / "template.md", root / "output",
                county="万州区",
            )
            engine.make_excel(
                config, county_segments,
                height_stats=height_stats, height_records=[record],
            )
            engine.add_charts(config.out_xlsx)
            workbook = openpyxl.load_workbook(config.out_xlsx)
            self.assertIn("二波分布图", workbook.sheetnames)
            charts = workbook["二波分布图"]._charts
            workbook.close()
        self.assertEqual(len(charts), 1)


class ChongqingMarkdownOnlyTests(unittest.TestCase):
    """R4：重庆链路只接受 Markdown 模板，不再读取 Word 模板。"""

    def test_docx_suffix_template_rejected_with_markdown_hint(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            fake = root / "template.docx"
            fake.write_bytes(b"fake")
            config = engine.Config(root, root / "summary.xlsx", root, fake, root / "output")
            with self.assertRaisesRegex(ValueError, "Markdown"):
                engine.make_docx(config, [], disease_image_index=None)

    def test_run_log_mentions_markdown_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            skeleton = root / "template.md"
            skeleton.write_text("# 标题\n", encoding="utf-8")
            config = engine.Config(root, root / "summary.xlsx", root, skeleton, root / "output")
            messages: list = []
            minimal_docx.run(
                config, [], None, [], None, [], None,
                log=messages.append, skeleton_md=skeleton,
            )
            self.assertTrue(any("Markdown" in message for message in messages))
            self.assertFalse(
                any("Word模板" in message or "Word 模板" in message for message in messages)
            )

    def test_chongqing_bridge_passes_markdown_template(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            project = root / "project"
            project.mkdir()
            detail = root / "detail"
            detail.mkdir()
            disease = root / "disease"
            disease.mkdir()
            tci = root / "tci"
            tci.mkdir()
            cq = root / "重庆项目报告模板.md"
            cq.write_text("# 重庆模板\n", encoding="utf-8")
            gd = root / "广东项目第五章模板.md"
            gd.write_text("# 广东模板\n", encoding="utf-8")
            bridge = DesktopBridge()
            captured: dict = {}

            def fake_run(config, **kwargs):
                captured["config"] = config

            values = {
                "projectPath": str(project),
                "summaryPath": str(root / "summary.xlsx"),
                "detailPath": str(detail),
                "diseasePath": str(disease),
                "tciPath": str(tci),
                "outputPath": str(root / "output"),
            }
            templates = {"重庆项目报告模板": cq, "广东项目第五章模板": gd}
            with patch.object(bridge, "_template_paths", return_value=templates), \
                    patch.object(engine, "generate_statistics_and_report", side_effect=fake_run):
                bridge._run_chongqing(values)
            self.assertEqual(captured["config"].template_docx.suffix, ".md")


class ChongqingSkeletonStructureTests(unittest.TestCase):
    """T10：重庆模板基准式显式锚点结构 + 只认锚点（无关键词回退）+ TCI 分段小节/占位。"""

    @staticmethod
    def _demo_tci_stats(segments):
        return engine.make_tci_stats(segments, [dict(segment=0, direction="上行", station=segments[0]["start"]+10,
                                                   light=1, heavy=0, sign=2, marking=10.)])

    def test_template_has_benchmark_anchor_structure(self) -> None:
        from backend import minimal_docx

        root = Path(__file__).resolve().parents[1]
        template = markdown_skeleton.read_template(root / "templates" / "重庆项目报告模板.md")
        h1 = [block.text for block in template.blocks if block.kind == "heading" and block.level == 1]
        self.assertEqual(
            [minimal_docx._strip_section_number(text) for text in h1],
            ["概况", "组织实施情况", "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测结果", "波形梁护栏螺栓缺失", "总结与建议"],
        )
        anchors = [block.text for block in template.blocks if block.text and "inject:" in block.text]
        self.assertEqual(len(anchors), 6)
        for key in ("overview", "tci", "height", "bolt", "conclusion", "tci_appendix"):
            self.assertEqual(sum(1 for text in anchors if f"inject:{key} -->" in text), 1, key)
        sub = [
            minimal_docx._strip_section_number(block.text)
            for block in template.blocks
            if block.kind == "heading" and block.level in (2, 3)
        ]
        for required in (
            "项目概况", "检测依据", "人员组织", "检测内容及方法", "检测设备与评定方法",
            "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测", "波形梁护栏螺栓缺失检测",
        ):
            self.assertIn(required, sub, required)
        for block in template.blocks:
            self.assertNotIn("两江新区", block.text or "")
            self.assertNotIn("由程序自动生成", block.text or "")

    def test_static_subsection_headings_are_not_swallowed_as_anchors(self) -> None:
        # 2.3.x 静态标题必须原样渲染，不得触发注入、不得被吞掉。
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            skeleton = root / "template.md"
            skeleton.write_text(
                "# 1 概况\n\n"
                "## 1.1 项目概况\n\n"
                "<!-- inject:overview -->\n\n"
                "## 1.2 检测依据\n\n"
                "依据文本。\n\n"
                "# 2 组织实施情况\n\n"
                "## 2.1 人员组织\n\n"
                "## 2.2 检测内容及方法\n\n"
                "## 2.3 检测设备与评定方法\n\n"
                "### 2.3.1 沿线设施技术状况评价\n\n"
                "评定方法文本。\n\n"
                "### 2.3.2 波形梁护栏横梁中心高度检测\n\n"
                "高度方法文本。\n\n"
                "### 2.3.3 波形梁护栏螺栓缺失检测\n\n"
                "螺栓方法文本。\n\n"
                "# 3 沿线设施技术状况评价\n\n"
                "<!-- inject:tci -->\n\n"
                "# 4 波形梁护栏横梁中心高度检测结果\n\n"
                "<!-- inject:height -->\n\n"
                "# 5 波形梁护栏螺栓缺失\n\n"
                "<!-- inject:bolt -->\n\n"
                "# 6 结论与建议\n\n"
                "<!-- inject:conclusion -->\n",
                encoding="utf-8",
            )
            config = engine.Config(root, root / "summary.xlsx", root, skeleton, root / "output")
            segments = build_demo_segments()
            records = build_demo_height_records()
            height_stats = build_demo_stats(segments, records)
            bolt_stats = engine.make_bolt_stats(segments, build_demo_bolt_records())
            tci_stats = self._demo_tci_stats(segments)
            result = minimal_docx.run(
                config, segments, height_stats, records, bolt_stats, build_demo_bolt_records(),
                None, skeleton_md=skeleton, tci_stats=tci_stats, tci_records=[],
            )
            document = Document(result)
            headings = [p.text for p in document.paragraphs if p.style.name.startswith("Heading")]
            text = "\n".join(p.text for p in document.paragraphs)
            h1 = [p.text for p in document.paragraphs if p.style.name == "Heading 1"]
            self.assertEqual(
                [minimal_docx._strip_section_number(h) for h in h1 if not h.startswith("附表1")],
                ["概况", "组织实施情况", "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测结果", "波形梁护栏螺栓缺失", "结论与建议"],
            )
        for required in ("检测依据", "检测设备与评定方法", "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测", "波形梁护栏螺栓缺失检测"):
            self.assertTrue(any(required in h for h in headings), f"missing static heading: {required}")
        # 按路线合并后，部分缺源不再建立独立区段小节。
        self.assertIn("部分检测区间未提供TCI数据", text)
        self.assertIn("G210线共检出拼接螺栓80颗", text)

    def test_tci_section_writes_route_subsection_and_excludes_missing_units(self) -> None:
        segments = build_demo_segments()
        tci_stats = self._demo_tci_stats(segments)
        document = Document()
        with tempfile.TemporaryDirectory() as temp_dir:
            tci_images = engine.report_tci_images(temp_dir, segments, tci_stats)
            image_files = {index: path.is_file() for index, path in tci_images.items()}
            minimal_docx._section_tci(document, segments, tci_stats, tci_images, temp_dir)
        headings = [p.text for p in document.paragraphs if p.style.name.startswith("Heading")]
        text = "\n".join(p.text for p in document.paragraphs)
        seg0 = f"G210线{engine.format_station(segments[0]['start'])}～{engine.format_station(segments[0]['end'])}段"
        seg1 = f"G210线{engine.format_station(segments[1]['start'])}～{engine.format_station(segments[1]['end'])}段"
        self.assertNotIn(seg0, headings)
        self.assertEqual(headings.count("G210线"), 1)
        self.assertNotIn(seg1, headings)
        self.assertIn("部分检测区间未提供TCI数据", text)
        self.assertEqual(len(tci_images), 1)
        self.assertTrue(all(image_files.values()), image_files)

    def test_tci_section_without_source_writes_explicit_sentence(self) -> None:
        document = Document()
        minimal_docx._section_tci(document, build_demo_segments(), None)
        text = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("未提供 TCI 病害清单", text)


class DiscoverPathsFallbackTests(unittest.TestCase):
    def _write_xlsx(self, path: Path, rows: list) -> None:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        for row in rows:
            sheet.append(row)
        workbook.save(path)

    def test_summary_fallback_matches整理_with_route_station_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            self._write_xlsx(
                root / "0. 各区县路段整理.xlsx",
                [["序号", "路线编号", "路线名", "起点桩号", "止点桩号", "里程"],
                 [1, "G210", "测试路", 1000.0, 2000.0, 1.0]],
            )
            summary, _, _, _ = engine.discover_paths(root)
        self.assertEqual(summary, root / "0. 各区县路段整理.xlsx")

    def test_summary_fallback_ignores整理_without_route_station_headers(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            self._write_xlsx(root / "人员整理.xlsx", [["姓名", "电话"], ["张三", "123"]])
            summary, _, _, _ = engine.discover_paths(root)
        self.assertIsNone(summary)

    def test_legacy_summary_keeps_priority_over_fallback(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            legacy = root / "G210采集路段信息汇总.xlsx"
            legacy.write_bytes(b"legacy")
            self._write_xlsx(
                root / "0. 各区县路段整理.xlsx",
                [["序号", "路线编号", "起点桩号"], [1, "G210", 1000.0]],
            )
            summary, _, _, _ = engine.discover_paths(root)
        self.assertEqual(summary, legacy)

    def test_detail_fallback_excludes_disease_and_requires_guardrail_header(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            detail_dir = root / "detail"
            detail_dir.mkdir()
            self._write_xlsx(
                detail_dir / "城口县-G211-明细.xlsx",
                [["序号", "路线编号", "护栏类型", "梁板中心高度(mm)"], [1, "G211", "二波", 600]],
            )
            disease_dir = root / "disease"
            disease_dir.mkdir()
            self._write_xlsx(
                disease_dir / "城口县-G211-病害明细.xlsx",
                [["序号", "路线", "病害类型"], [1, "G211", "螺栓缺失"]],
            )
            plain_dir = root / "plain"
            plain_dir.mkdir()
            self._write_xlsx(plain_dir / "说明-明细.xlsx", [["姓名"], ["张三"]])
            _, detail, disease, _ = engine.discover_paths(root)
        self.assertEqual(detail, detail_dir)
        self.assertEqual(disease, disease_dir)


class T11aDataLayerTests(unittest.TestCase):
    """T11a：螺栓 sheet 读取 + 浮动图片锚点映射（含装饰图跳过）。"""

    @staticmethod
    def _write_xlsx(path, rows):
        import openpyxl as xl

        wb = xl.Workbook()
        ws = wb.active
        for row in rows:
            ws.append(row)
        wb.save(path)
        wb.close()

    @staticmethod
    def _tiny_png() -> bytes:
        import base64

        return base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )

    def test_iter_bolt_rows_finds_sheet_via_shared_string_index(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "明细.xlsx"
            self._write_xlsx(
                path,
                [["序号", "路线编号", "梁板中心高度(mm)"], [1, "G211", 600.0]],
            )
            wb = openpyxl.load_workbook(path)
            ws2 = wb.create_sheet("螺栓缺失")
            ws2.append(["序号", "路线编号", "电子修正桩号", "拼接螺栓数量（颗）", "拼接螺栓缺失数量（颗）", "连接螺栓数量（颗）", "连接螺栓缺失数量（颗）"])
            ws2.append([1, "G211", "1157+874", 10, 1, 20, 2])
            ws2.append([2, "G211", "1158+000", 0, 0, 0, 0])
            wb.save(path)
            wb.close()
            rows = list(engine.iter_bolt_rows(path))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["拼接螺栓数量（颗）"], 10)
        self.assertEqual(rows[0]["连接螺栓缺失数量（颗）"], 2)

    def test_iter_bolt_rows_ignores_sheets_without_marker(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "明细.xlsx"
            self._write_xlsx(
                path,
                [["序号", "路线编号", "梁板中心高度(mm)"], [1, "G211", 600.0]],
            )
            rows = list(engine.iter_bolt_rows(path))
        self.assertEqual(rows, [])

    def test_build_disease_image_map_skips_header_decorations(self) -> None:
        import io

        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "病害.xlsx"
            self._write_xlsx(
                path,
                [
                    ["序号", "路线", "原始桩号", "病害类型", "工程量", "单位", "病害照片", "备注"],
                    [1, "G211", "1157+874", "高度异常", 1, "处", None, None],
                    [2, "G211", "1158+000", "螺栓缺失", 1, "颗", None, None],
                ],
            )
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            png = self._tiny_png()
            ws.add_image(XLImage(io.BytesIO(png)), "A1")  # 表头行装饰图，应跳过
            ws.add_image(XLImage(io.BytesIO(png)), "M2")  # 第一数据行照片列(0-based col12)
            ws.add_image(XLImage(io.BytesIO(png)), "M3")
            wb.save(path)
            wb.close()
            mapping = engine.build_disease_image_map(path)
        rows_with_images = {row for (sheet, row) in mapping}
        self.assertEqual(rows_with_images, {2, 3})
        self.assertTrue(all(sheet.endswith(".xml") for (sheet, _) in mapping))
        self.assertTrue(all(len(v) == 1 for v in mapping.values()))

    def test_build_tci_image_map_uses_photo_column_15(self) -> None:
        import io

        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "tci.xlsx"
            self._write_xlsx(
                path,
                [
                    ["区域", "路线编号", "原始桩号", "防护设施缺损（处）", "标志缺损（处）", "标线缺损（m）", "图片"],
                    ["万州区", "G348", "949+680", 1, 0, 0, None],
                    ["万州区", "G348", "950+000", 0, 1, 5, None],
                ],
            )
            wb = openpyxl.load_workbook(path)
            ws = wb.active
            png = self._tiny_png()
            ws.add_image(XLImage(io.BytesIO(png)), "P3")  # 0-based col15, row2
            wb.save(path)
            wb.close()
            mapping = engine.build_tci_image_map(path)
        self.assertEqual(set(mapping.keys()), {("xl/worksheets/sheet1.xml", 3)})

    def test_height_collection_selects_height_sheet_and_keeps_route_identity(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "multi-sheet-detail.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "说明"
            ws.append(["说明", "值"])
            ws.append(["不是护栏明细", 1])
            ws = wb.create_sheet("高度明细")
            ws.append(["备注", "路线编号", "电子修正桩号", "护栏类型", "梁板中心高度(mm)"])
            ws.append([None, "G210", "K1+100", "二波", 600])
            ws = wb.create_sheet("螺栓明细")
            ws.append(["路线编号", "电子修正桩号", "拼接螺栓数量（颗）", "拼接螺栓缺失数量（颗）", "连接螺栓数量（颗）", "连接螺栓缺失数量（颗）"])
            ws.append(["G210", "K1+100", 10, 1, 10, 0])
            wb.save(path)
            wb.close()
            segments = [{"county": "甲县", "route": "G210", "start": 1000, "end": 2000, "mileage": 1.0}]
            records, duplicates, excluded = engine.collect_records(segments, root)
        self.assertEqual(duplicates, 0)
        self.assertEqual(excluded, {})
        self.assertEqual(len(records), 1)
        self.assertEqual(records[0]["station"], 1100)
        self.assertEqual(records[0]["route"], "G210")
        self.assertEqual(records[0]["county"], "甲县")
        self.assertEqual(records[0]["source_sheet"], "高度明细")
        self.assertEqual(records[0]["source_row"], 2)

    def test_bolt_collection_filters_only_blank_or_no_remark_marker(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "bolt-detail.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "螺栓明细"
            ws.append(["路线编号", "电子修正桩号", "备注标记", "拼接螺栓数量（颗）", "拼接螺栓缺失数量（颗）", "连接螺栓数量（颗）", "连接螺栓缺失数量（颗）"])
            for index, remark in enumerate([None, "", "   ", "无备注", " 无备注 ", "待复核"], 1):
                ws.append(["G210", f"K1+{index:03d}", remark, 10, 1, 10, 0])
            wb.save(path)
            wb.close()
            segments = [{"county": "甲县", "route": "G210", "start": 1000, "end": 2000, "mileage": 1.0}]
            records, duplicates = engine.collect_bolt_records(segments, root)
            stats = engine.make_bolt_stats(segments, records)
        self.assertEqual(duplicates, 0)
        self.assertEqual(len(records), 5)
        self.assertEqual(stats[0]["splice"], 50)
        self.assertEqual(stats[0]["missing"], 5)
        self.assertAlmostEqual(stats[0]["rate"], 5 / 105 * 100)

    def test_photo_column_can_move_and_height_does_not_match_other_route_or_bolt(self) -> None:
        import io
        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "photos.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "无关"
            ws.append(["说明"])
            ws = wb.create_sheet("照片明细")
            ws.append(["路线编号", "方向", "原始桩号", "病害类型", "工程量", "备注", "病害照片"])
            ws.append(["G210", "上行", "K1+100", "高度异常", 1, None, None])
            ws.append(["G211", "上行", "K1+100", "高度异常", 1, None, None])
            ws.append(["G210", "上行", "K1+101", "螺栓缺失", 1, None, None])
            png = self._tiny_png()
            ws.add_image(XLImage(io.BytesIO(png)), "G2")
            ws.add_image(XLImage(io.BytesIO(png)), "G3")
            ws.add_image(XLImage(io.BytesIO(png)), "G4")
            wb.save(path)
            wb.close()
            image_map = engine.build_disease_image_map(path)
            index = engine.disease_station_photo_index(path, image_map, role="height")
            good = {"route": "G210", "county": "甲县", "direction": "上行", "station": 1100, "raw_station": 1100, "electronic_station": 1100}
            wrong_route = dict(good, route="G212")
        self.assertEqual(set(image_map), {("xl/worksheets/sheet2.xml", 2), ("xl/worksheets/sheet2.xml", 3), ("xl/worksheets/sheet2.xml", 4)})
        self.assertTrue(engine.row_has_height_photo(good, index))
        self.assertFalse(engine.row_has_height_photo(wrong_route, index))
        self.assertFalse(engine.row_has_height_photo(dict(good, station=1101, raw_station=1101, electronic_station=1101), index))

    def test_disease_image_index_keeps_workbook_identity_when_media_names_repeat(self) -> None:
        import io
        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            disease_dir = root / "disease"
            disease_dir.mkdir()
            for name in ("first.xlsx", "second.xlsx"):
                path = disease_dir / name
                wb = openpyxl.Workbook()
                ws = wb.active
                ws.append(["路线编号", "方向", "原始桩号", "病害类型", "工程量", "病害照片"])
                ws.append(["G210", "上行", "K1+100", "螺栓缺失", 2, None])
                ws.add_image(XLImage(io.BytesIO(self._tiny_png())), "F2")
                wb.save(path)
                wb.close()
            image_index = engine.collect_disease_image_index(disease_dir)
            record = {"route": "G210", "county": "甲县", "direction": "上行", "raw_station": 1100,
                      "splice_missing": 2, "connection_missing": 0}
            descriptor = engine.match_disease_image(record, image_index)
        self.assertIsNotNone(descriptor)
        self.assertEqual({item["workbook"].name for item in image_index[("上行", 1100.0)]}, {"first.xlsx", "second.xlsx"})
        self.assertIn(descriptor["workbook"].name, {"first.xlsx", "second.xlsx"})

    def test_disease_image_index_parses_each_sheet_once(self) -> None:
        """照片行不得触发整表重复解析：真实病害清单 4314 行照片曾导致 4314 次全表解析（约 54 分钟）。"""
        import io
        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            disease_dir = root / "disease"
            disease_dir.mkdir()
            path = disease_dir / "many.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["路线编号", "方向", "原始桩号", "病害类型", "工程量", "病害照片"])
            for index in range(12):
                ws.append(["G210", "上行", f"K{1 + index}+100", "螺栓缺失", 2, None])
                ws.add_image(XLImage(io.BytesIO(self._tiny_png())), f"F{index + 2}")
            wb.save(path)
            wb.close()
            calls: list = []
            original = engine._xlsx_sheet_rows

            def counting(archive, sheet_name, shared_strings):
                calls.append(sheet_name)
                return original(archive, sheet_name, shared_strings)

            engine._xlsx_sheet_rows = counting
            try:
                engine.collect_disease_image_index(disease_dir)
            finally:
                engine._xlsx_sheet_rows = original
        self.assertEqual(len(calls), len(set(calls)), f"同一 sheet 被重复解析 {len(calls)} 次：{calls}")

    def test_tci_photo_index_reads_nonfirst_sheet_and_keeps_route_filter(self) -> None:
        import io
        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            path = root / "tci-multi-sheet.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "说明"
            ws.append(["说明"])
            ws = wb.create_sheet("TCI数据")
            headers = ["区域", "路线编号", "方向", "原始桩号", "防护设施缺损（处）", "标志缺损（处）", "标线缺损（m）", "防护设施缺损", "交通标志缺损", "交通标线缺损", "图片"]
            ws.append(headers)
            ws.append(["甲县", "G210", "上行", "K1+100", 1, 0, 0, "护栏锈蚀", "", "", None])
            ws.append(["甲县", "G211", "上行", "K1+200", 0, 1, 0, "", "标志遮挡", "", None])
            png = self._tiny_png()
            ws.add_image(XLImage(io.BytesIO(png)), "K2")
            ws.add_image(XLImage(io.BytesIO(png)), "K3")
            wb.save(path)
            wb.close()
            image_map = engine.build_tci_image_map(path)
            filtered = engine.tci_type_photo_index(path, image_map, routes=["G210"])
            segments = [{"county": "甲县", "route": "G210", "start": 1000, "end": 1500, "mileage": .5}]
            segment_photos = engine.build_segment_tci_photos(path, image_map, segments)
        self.assertEqual(set(image_map), {("xl/worksheets/sheet2.xml", 2), ("xl/worksheets/sheet2.xml", 3)})
        self.assertIn("防护设施缺损", filtered)
        self.assertNotIn("交通标志缺损", filtered)
        self.assertEqual(set(segment_photos), {0})

    def test_overview_sums_segment_mileage_and_merges_total_by_county_route(self) -> None:
        segments = [
            {"county": "甲县", "route": "G210", "route_name": "甲", "grade": "一级", "start": 1000, "end": 1100, "mileage": .1, "total_mileage": 99},
            {"county": "甲县", "route": "G211", "route_name": "乙", "grade": "一级", "start": 2000, "end": 2200, "mileage": .2, "total_mileage": 88},
            {"county": "甲县", "route": "G210", "route_name": "甲", "grade": "一级", "start": 1100, "end": 1300, "mileage": .2, "total_mileage": 77},
            {"county": "乙县", "route": "G210", "route_name": "丙", "grade": "一级", "start": 3000, "end": 3500, "mileage": .5, "total_mileage": 66},
        ]
        document = Document()
        table = minimal_docx._section_overview(document, None, segments)
        rows = [[cell.text for cell in row.cells] for row in table.rows]
        total_col = len(rows[0]) - 1
        xml = table._tbl.xml
        self.assertEqual([rows[i][0] for i in range(1, 5)], ["1", "3", "2", "4"])
        self.assertEqual([rows[i][total_col] for i in range(1, 5)], ["0.300", "0.300", "0.200", "0.500"])
        self.assertEqual(xml.count('w:val="restart"'), 1)
        self.assertGreaterEqual(xml.count('w:val="continue"'), 1)


class T11bPresentationTests(unittest.TestCase):
    """T11b：封面黄高亮/黄块/规则页眉线 + TCI 公式 OMML。"""

    _TEMPLATE = (
        "# @cover 主标题行一|主标题行二|重庆市|报告编号：BG-2026-T11"
        "|项目名称：测试项目|委托单位：测试单位|测试公司|二〇二六年七月\n"
        "\n# @notes 注意事项|1、测试条款|6、联系方式：\n"
        "\n<!-- toc -->\n\n# 1 概况\n## 1.1 项目概况\n<!-- inject:overview -->\n"
        "# 2 组织实施情况\n## 2.3 检测设备与评定方法\n### 2.3.1 沿线设施技术状况评价\n"
        "<!-- formula:tci -->\n"
        "<!-- formula:gd --> —— 第i类设施损坏的累计扣分，最高扣分为100，按表2.3.1-1的规定取值；\n"
        "# 3 沿线设施技术状况评价\n<!-- inject:tci -->\n"
    )

    def _build(self) -> dict:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            template = root / "template.md"
            template.write_text(self._TEMPLATE, encoding="utf-8")
            output = root / "output"
            output.mkdir()
            config = engine.Config(root, root / "summary.xlsx", root, template, output)
            minimal_docx.make_report(
                config, [], None, [], None, [], None, [], {}, root,
                skeleton_md=template,
            )
            out = output / engine.OUT_DOCX_NAME
            with zipfile.ZipFile(out) as archive:
                parts = {
                    name: archive.read(name).decode("utf-8")
                    for name in archive.namelist()
                    if name.startswith("word/header") or name == "word/document.xml"
                }
            return parts

    def test_cover_highlight_and_yellow_block_present(self) -> None:
        parts = self._build()
        document = parts["word/document.xml"]
        self.assertIn("w:highlight", document)
        self.assertIn('w:val="yellow"', document)
        # 编号值 BG-2026-T11 与项目名称值都被高亮
        self.assertIn("BG-2026-T11", document)
        self.assertIn("测试项目", document)

    def test_tci_formula_is_editable_omath_without_plaintext_duplicate(self) -> None:
        parts = self._build()
        document = parts["word/document.xml"]
        self.assertGreaterEqual(document.count("m:oMath"), 2)
        self.assertIn("TCI=", document)
        self.assertNotIn("TCI = Σ_{i=1}", document)
        self.assertIn("第i类设施损坏的累计扣分", document)

    def test_cover_section_header_has_rule(self) -> None:
        parts = self._build()
        headers = [xml for name, xml in parts.items() if "/header" in name]
        self.assertTrue(headers, "应生成页眉部件")
        self.assertTrue(any("w:pBdr" in xml and "w:bottom" in xml for xml in headers))


class R12TemplateTests(unittest.TestCase):
    """R12/R13：区间分档 + 示例图统一15×8cm无锁 + TCI类型图路线过滤。"""

    def test_height_limits_new_bands(self) -> None:
        from backend.report_engine import bin_index

        # 二波：550/580/620/650（中心600，±20合格，±50外界）
        self.assertEqual(bin_index("二波", 545.0), 0)
        self.assertEqual(bin_index("二波", 555.0), 1)
        self.assertEqual(bin_index("二波", 600.0), 2)
        self.assertEqual(bin_index("二波", 625.0), 3)
        self.assertEqual(bin_index("二波", 655.0), 4)
        # 三波：647/677/717/747（中心697）
        self.assertEqual(bin_index("三波", 640.0), 0)
        self.assertEqual(bin_index("三波", 650.0), 1)
        self.assertEqual(bin_index("三波", 700.0), 2)
        self.assertEqual(bin_index("三波", 720.0), 3)
        self.assertEqual(bin_index("三波", 750.0), 4)

    def test_example_photo_table_15x8_unlocked_and_tables_separated_only_after_data_table(self) -> None:
        import tempfile

        from docx import Document

        from backend import minimal_docx

        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
        with tempfile.TemporaryDirectory() as tmp:
            # 1) 多列结果表之后必须插入独立段落（防与示例表合并）
            doc = Document()
            minimal_docx._table(doc, ["起点桩号", "止点桩号", "缺失数量（颗）"], [["K1", "K2", "3"]], True)
            minimal_docx._photo_text_table(doc, png, ".png", "标志板遮挡", tmp, "test_photo")
            out = Path(tmp) / "t.docx"
            doc.save(out)
            with zipfile.ZipFile(out) as z:
                xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn('cx="5400000"', xml)  # 15cm
            self.assertIn('cy="2880000"', xml)  # 8cm
            self.assertIn('noChangeAspect="0"', xml)
            self.assertRegex(xml, r"</w:tbl><w:p[ >].*?</w:p><w:tbl>")
            self.assertIn("标志板遮挡", xml)
        with tempfile.TemporaryDirectory() as tmp:
            # 2) 单列示例表之间不加空段（两张示例图片连续排版）
            doc = Document()
            minimal_docx._photo_text_table(doc, png, ".png", "示例一", tmp, "t1")
            minimal_docx._photo_text_table(doc, png, ".png", "示例二", tmp, "t2")
            out = Path(tmp) / "t.docx"
            doc.save(out)
            with zipfile.ZipFile(out) as z:
                xml = z.read("word/document.xml").decode("utf-8")
            self.assertNotRegex(xml, r"</w:tbl><w:p[ >].*?</w:p><w:tbl>")
            self.assertRegex(xml, r"</w:tbl><w:tbl>")
            self.assertIn("示例一", xml)
            self.assertIn("示例二", xml)

    def test_tci_type_photo_index_routes_filter(self) -> None:
        import openpyxl

        from backend import report_engine as engine

        png = base64.b64decode(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
        )
        with tempfile.TemporaryDirectory() as tmp:
            tmp_path = Path(tmp)
            xlsx = tmp_path / "tci.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            headers = ["区域", "路线编号", "方向", "原始桩号", "标注修正桩号", "电子修正桩号", "经度", "纬度",
                       "防护设施缺损（处）", "J", "标志缺损（处）", "标线缺损（m）", "防护设施缺损", "交通标志缺损", "交通标线缺损"]
            ws.append(headers)
            ws.append(["重庆市", "G348", "上行", "K949+680", "K949+680", "K949+680", "", "", 0, "重", 1, 0, "", "标志板遮挡", ""])
            ws.append(["重庆市", "G210", "上行", "K1157+874", "K1157+874", "K1157+874", "", "", 1, "重", 0, 0, "波形梁护栏锈蚀", "", ""])
            img = openpyxl.drawing.image.Image(io.BytesIO(png))
            img.width = 16
            img.height = 8.5
            ws.add_image(img, "P2")
            ws.add_image(openpyxl.drawing.image.Image(io.BytesIO(png)), "P3")
            wb.save(xlsx)
            image_map = engine.build_tci_image_map(xlsx)
            self.assertTrue(image_map, "合成 TCI 图片映射为空")
            filtered = engine.tci_type_photo_index(xlsx, image_map, routes=["G348"])
            self.assertIn("交通标志缺损", filtered)
            self.assertNotIn("防护设施缺损", filtered)
            all_types = engine.tci_type_photo_index(xlsx, image_map)
            self.assertIn("防护设施缺损", all_types)


if __name__ == "__main__":
    unittest.main()
