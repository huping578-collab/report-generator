from __future__ import annotations

import base64
import io
import json
import os
import openpyxl
import re
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


def test_chongqing_three_chapters_preserve_sorted_section_direction(tmp_path):
    segments = [dict(county='甲县', route='G210', start=a, end=b, mileage=(b-a)/1000)
                for a, b in [(5000, 6000), (1000, 2000)]]
    heights = [dict(segment=i, route='G210', county='甲县', kind='二波', height=600,
                    direction=d, station=s['start']+10, raw_station=s['start']+10,
                    electronic_station=s['start']+10)
               for i, s in enumerate(segments) for d in engine.DIRECTIONS]
    bolts = [dict(r, splice=10, connection=10, splice_missing=1, connection_missing=0) for r in heights]
    tci = [dict(r, light=1, heavy=0, sign=0, marking=0) for r in heights]
    hs, bs, ts = engine.make_stats(segments, heights), engine.make_bolt_stats(segments, bolts), engine.make_tci_stats(segments, tci)
    images = engine.report_images(tmp_path, segments, hs, heights)
    for chapter in ('height', 'bolt', 'tci'):
        doc = Document()
        if chapter == 'height':
            minimal_docx._section_height(doc, segments, hs, heights, images, tmp_path)
            assert len(doc.inline_shapes) == 8
            assert all('检测里程（km）' not in [c.text for c in t.rows[0].cells] for t in doc.tables)
        elif chapter == 'bolt':
            minimal_docx._section_bolt(doc, segments, bs, bolts, {}, tmp_path, hs)
        else:
            minimal_docx._section_tci(doc, segments, ts, engine.report_tci_images(tmp_path, segments, ts), tmp_path)
        headings = [p.text for p in doc.paragraphs if p.style.name == 'Heading 3']
        assert headings == [f'G210线{d}{engine.format_station(a)}～{engine.format_station(b)}段'
                            for a,b in [(1000,2000),(5000,6000)] for d in engine.DIRECTIONS]
        captions = [p.text.split()[0] for p in doc.paragraphs if p.text.startswith(('图4.', '图5.', '图3.'))]
        assert len(captions) == len(set(captions))
        route_tables = [t for t in doc.tables if '方向' in [c.text for c in t.rows[0].cells]]
        assert route_tables and all('起点桩号' in [c.text for c in t.rows[0].cells] for t in route_tables)


def test_height_photo_consumer_filters_each_candidate(tmp_path):
    from PIL import Image
    point = dict(route='G210', county='甲县', direction='上行', station=1100,
                 raw_station=1100, electronic_station=1100, height=600)
    candidates = []
    for index, identity in enumerate([dict(route='G211'), dict(county='乙县'), dict(role='bolt'), {}]):
        descriptor = dict(point, role='height', workbook=tmp_path / f'{index}.xlsx', media='xl/media/image1.png')
        descriptor.update(identity)
        data = io.BytesIO()
        Image.new('RGB', (10,10), (index*50,0,0)).save(data, format='PNG')
        with zipfile.ZipFile(descriptor['workbook'], 'w') as archive:
            archive.writestr(descriptor['media'], data.getvalue())
        candidates.append(descriptor)
    result = minimal_docx._match_station_photos([point], {('上行',1100):candidates})
    expected = engine.read_media(candidates[-1]['workbook'], candidates[-1]['media'])
    assert result == {0:[expected]}
    doc = Document()
    minimal_docx._example_table(doc, [point], result, tmp_path)
    assert [(shape.width, shape.height) for shape in doc.inline_shapes] == [(2160000,2880000)]
    assert 'noChangeAspect="1"' not in doc._element.xml


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
    # 重庆口径：方向按源文件名判定（identity.xlsx 不含“下行”），表内方向仅用于冲突提示。
    assert [r["segment"] for r in records] == [1, 0, 3, 4]
    assert [r["direction"] for r in records] == ["上行", "上行", "上行", "上行"]
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
    # 附表：同一区县一张表，路线分块——每条路线前重出表头、序号从 1 重排，块间整行合并空白行分隔。
    assert '附表1 重庆市甲县交安设施技术状况评定明细' in text
    assert '评定单元等权' in text
    assert '共检出拼接螺栓20颗' in text
    appendix_tables = [table for table in doc.tables if [c.text for c in table.rows[0].cells] ==
                       ['序号','路线编号','区县','方向','起点桩号','止点桩号','TCI','等级']]
    assert len(appendix_tables) == 1  # G210 与 G319 同表分块
    appendix = appendix_tables[0]
    assert appendix.rows[0]._tr.xpath('./w:trPr/w:tblHeader')
    assert [row.cells[0].text for row in appendix.rows] == ['序号','1','2','3','4','5','','序号','1','2']
    assert appendix.rows[1].cells[4].text == '/'  # 路线汇总行的起止桩号留空
    assert appendix.rows[7].cells[0].paragraphs[0].runs[0].bold is True  # 重复表头按表头样式
    separator = appendix.rows[6]._tr.findall(qn('w:tc'))
    assert len(separator) == 1 and separator[0].find(qn('w:tcPr')).find(qn('w:gridSpan')).get(qn('w:val')) == '8'
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
        assert axis.lines[0].get_color() == engine.GD_SERIES_HEX[0]   # P1c：模板 2007-2010 accent1
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


class GuangdongStructureAlignmentTests(unittest.TestCase):
    """第五章结构对齐附件模板（交通安全设施技术状况检测评价报告模板-2.docx）：
    2.标线、护栏自动化检测 下三个指标组均为（n）标题 + ①②各抽检路段情况③典型不佳 + ④长连续梳理，
    且每组①总体情况带“全省/本市”两行汇总表；模板自带的 HTML 注释不得进入正文。"""

    @staticmethod
    def _bundle() -> dict:
        def height(kind, value, seg="K1+000～K2+000"):
            return {"category": "高速公路", "route": "S14", "direction": "上行", "segment": seg,
                    "guardrail_type": kind, "height": value, "city": "佛山市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 2000.0}
        return {
            "city": "佛山市",
            "marking": [{"category": "高速公路", "route": "S14", "direction": "上行",
                         "segment": "K1+000～K2+000", "city": "佛山市",
                         "manager": "广东省高速公路有限公司", "position_name": "左侧标线",
                         "station_m": 1000.0, "end_m": 2000.0, "target": 80.0, "coefficient": 90.0}],
            "height": [height("二波", 600), height("三波", 700)],
            "bolt": [{"category": "高速公路", "route": "S14", "direction": "上行",
                      "segment": "K1+000～K2+000", "city": "佛山市",
                      "manager": "广东省高速公路有限公司", "guardrail_type": "二波",
                      "splice": 100, "splice_missing": 2, "connection": 50, "connection_missing": 1,
                      "station_m": 1000.0, "end_m": 2000.0}],
            "notes": [], "comparison_detail": [], "comparison_summary": {}, "weak_segments": [],
        }

    def test_group_headings_match_attachment_and_notes_are_stripped(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "佛山市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            texts = [p.text.strip() for p in document.paragraphs]

        # G2（用户第 3 轮裁决，覆盖 D7/R10）：三个自动化小节标题统一带 `（1）（2）（3）`
        self.assertIn("（1）标线逆反射亮度系数情况", texts)
        self.assertIn("（2）波形梁护栏中心高度情况", texts)
        self.assertIn("（3）波形梁护栏螺栓缺失情况", texts)
        self.assertNotIn("（2）护栏中心高度情况", texts)
        # 三个指标组各自的①②小节（附件口径）
        self.assertEqual(texts.count("②各抽检路段情况"), 2)  # 本样例仅有高度/螺栓数据
        # 模板注释不得进入正文
        self.assertFalse([t for t in texts if t.startswith("<!--")])

    def test_overall_tables_use_manager_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "佛山市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 7, "height": 5, "bolt": 5},
            )
            document = Document(output)
            headers = [[c.text.strip() for c in t.rows[0].cells] for t in document.tables]
            first_cols = [t.rows[i].cells[0].text.strip()
                          for t in document.tables for i in range(1, len(t.rows))]

        # brief-v3 §1 / P1b：高速 ① 表为「管理单位」级 5 列，逐字 4 行（全省抽检高速 / XX市抽检高速 /
        # 非省交通集团 / 省交通集团）；第 1 行是全省行，4 个数据格全部留空。
        overall = [h for h in headers if h and h[0] == "管理单位"]
        self.assertTrue(overall, f"未找到①总体情况汇总表（管理单位级），现有表头：{headers}")
        self.assertTrue(all(len(h) == 5 for h in overall), f"①表应为 5 列（模板形态）：{overall}")
        for t in (t for t in document.tables if t.rows[0].cells[0].text.strip() == "管理单位"):
            labels = [t.rows[i].cells[0].text.strip() for i in range(1, len(t.rows))]
            self.assertEqual(labels, ["全省抽检高速", "佛山市抽检高速", "非省交通集团", "省交通集团"])
            self.assertEqual([c.text.strip() for c in t.rows[1].cells[1:]], ["", "", "", ""],
                             "全省抽检高速行的 4 个数据格必须留空，不填占位也不填「—」")


class T2bTemplateAlignmentTests(unittest.TestCase):
    """T2b：① 汇总表（D8）、抽检路段清单表（D2/D3/D19/D4）、② 前置句（D10）按甲方模板逐字对齐。"""

    @staticmethod
    def _bundle() -> dict:
        def height(value, station):
            return {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K2+000",
                    "guardrail_type": "二波", "height": value, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": station, "end_m": station + 1}

        return {
            "city": "测试市",
            "marking": [],
            "height": [height(600.0, 1000.0), height(600.0, 2000.0)],
            "bolt": [],
            "notes": [],
            "comparison_detail": [],
            "weak_segments": [],
            "route_segments": [
                {"city": "测试", "route": "S14", "direction": "上行", "category": "高速公路",
                 "start": 964.986, "end": 1009.0, "length": 44.014, "manager": "广东省高速公路有限公司"},
                {"city": "测试", "route": "S14", "direction": "下行", "category": "高速公路",
                 "start": 1009.0, "end": 2000.0, "length": 991.0, "manager": "广东省高速公路有限公司"},
                {"city": "测试", "route": "S357", "direction": "上行", "category": "普通国省道",
                 "start": 40.0109, "end": 60.0, "length": 19.9891, "manager": "测试市公路事务中心"},
                {"city": "测试", "route": "S357", "direction": "下行", "category": "普通国省道",
                 "start": 60.0, "end": 80.0, "length": 20.0, "manager": "测试市公路事务中心"},
            ],
        }

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_inspection_list_table_matches_template(self) -> None:
        document = self._write()
        captions = [p.text for p in document.paragraphs if p.style.name == "Caption"]
        table = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "类型")
        headers = [c.text.strip() for c in table.rows[0].cells]

        # D2 表题去「表」；D3 第 7 列 `段长（Km)`；D5 声明在 GD_TABLE_WIDTHS
        self.assertIn("表5-1 测试市交通安全设施抽检路段清单", captions)
        self.assertEqual(headers, ["类型", "路线编号", "路线名称", "检测方向", "起点桩号", "终点桩号",
                                   "段长（Km)", "管养单位", "备注"])
        self.assertIsNotNone(engine.GD_TABLE_WIDTHS.get(tuple(engine._norm_header_text(h) for h in headers)))
        # Q3 整数公里：964.986→965、40.0109→40、44.014→44
        self.assertEqual([table.rows[1].cells[i].text.strip() for i in (4, 5, 6)], ["965", "1009", "44"])
        self.assertEqual([table.rows[2].cells[i].text.strip() for i in (4, 5, 6)], ["1009", "2000", "991"])
        self.assertEqual([table.rows[3].cells[i].text.strip() for i in (4, 5, 6)], ["40", "60", "20"])

        # D4 类型列按高速/普通块纵向合并（连续同值合并，首行 restart）
        xml_rows = table._tbl.findall("./w:tr", table._tbl.nsmap)

        def vmerge_of(row):
            cell = row.findall("./w:tc", table._tbl.nsmap)[0]
            node = cell.find("./w:tcPr/w:vMerge", cell.nsmap)
            return "none" if node is None else (node.get(qn("w:val")) or "continue")

        self.assertEqual([vmerge_of(row) for row in xml_rows],
                         ["none", "restart", "continue", "restart", "continue"])

    def test_overall_table_and_preface_sentence_match_template(self) -> None:
        document = self._write()
        overall = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "管理单位")
        body = "\n".join(p.text for p in document.paragraphs)

        # brief-v3 §1 / P1b：高速 ① 表逐字 4 行，首行「全省抽检高速」且数据格留空
        self.assertEqual([c.text.strip() for c in overall.rows[0].cells],
                         ["管理单位", "抽检里程(km)", "总体合格率(%)", "两波合格率(%)", "三波合格率(%)"])
        self.assertEqual([overall.rows[i].cells[0].text.strip() for i in range(1, len(overall.rows))],
                         ["全省抽检高速", "测试市抽检高速", "非省交通集团", "省交通集团"])
        self.assertEqual([c.text.strip() for c in overall.rows[1].cells[1:]], ["", "", "", ""])
        # D10：② 前置句逐字（模板句型）
        self.assertIn("抽检的各高速公路路段波形梁护栏中心高度总体合格率明细如下表所示。", body)


class T2dTableLayoutTests(unittest.TestCase):
    """T2d：列宽取自 GD_TABLE_WIDTHS(_BY_CAPTION) 且与 template-tables.tsv 一致；② 表尾行合计合并。"""

    @staticmethod
    def _bundle() -> dict:
        rows = [
            {"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K2+000",
             "guardrail_type": kind, "height": 600.0, "city": "测试市", "manager": "广东省高速公路有限公司",
             "station_m": station, "end_m": station + 1.0}
            for station, kind in ((1000.0, "两波"), (1001.0, "三波"))
        ]
        return {"city": "测试市", "marking": [], "bolt": [], "notes": [], "comparison_detail": [],
                "weak_segments": [], "height": rows,
                "route_segments": [{"city": "测试", "route": "S14", "direction": "上行", "category": "高速公路",
                                    "start": 1.0, "end": 2.0, "length": 1.0, "manager": "广东省高速公路有限公司"}]}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_widths_dict_matches_template_tsv(self) -> None:
        # 清单表 / 人工复核 / 分支列宽不同（表题口径）：数值取自 requirements/template-tables.tsv
        self.assertEqual(
            engine.GD_TABLE_WIDTHS[("类型", "路线编号", "路线名称", "检测方向", "起点桩号", "终点桩号",
                                    "段长（Km)", "管养单位", "备注")],
            [1.36, 1.04, 1.78, 1.09, 1.09, 1.12, 1.31, 4.25, 2.47])
        self.assertEqual(
            engine.GD_TABLE_WIDTHS[("路线", "桩号区段", "护栏类型", "人工检测复核结果", "人工检测复核结果",
                                    "人工检测复核结果", "自动化检测结果", "自动化检测结果", "自动化检测结果", "结果一致性")],
            [1.52, 2.56, 1.26, 1.31, 1.31, 1.32, 1.32, 1.32, 1.32, 1.37])
        self.assertEqual(engine.GD_TABLE_WIDTHS_BY_CAPTION["普通国省道标线逆反射亮度系数不佳路段汇总表"],
                         [1.13, 5.91, 2.21, 1.85, 1.79, 1.80])
        # 人工复核三表（标线两分支同宽；高度/螺栓两分支列宽不同 → 表题口径）
        self.assertEqual(
            engine.GD_TABLE_WIDTHS[("路线编号", "桩号区段", "标线颜色", "人工检测复核结果", "人工检测复核结果",
                                    "自动化检测结果", "自动化检测结果", "结果一致性")],
            [1.34, 3.55, 1.63, 1.63, 1.79, 1.46, 1.76, 1.58])
        self.assertEqual(
            engine.GD_TABLE_WIDTHS[("路线编号", "桩号区段", "护栏类型", "合格值(mm)", "人工检测复核结果",
                                    "人工检测复核结果", "自动化检测结果", "自动化检测结果", "结果一致性")],
            [1.23, 3.26, 1.0, 2.17, 1.25, 1.59, 1.23, 1.62, 1.44])
        self.assertEqual(engine.GD_TABLE_WIDTHS_BY_CAPTION["普通国省道波形梁护栏中心高度人工复核对比明细表"],
                         [1.23, 3.29, 1.0, 2.16, 1.25, 1.59, 1.23, 1.62, 1.43])

    def test_route_table_widths_and_tail_total_row(self) -> None:
        document = self._write()
        table = next(t for t in document.tables
                     if len(t.columns) == 8 and t.rows[0].cells[4].text.strip() == "里程(km)")
        self.assertEqual([round(column.width.cm, 2) for column in table.columns],
                         [1.01, 1.60, 3.64, 2.21, 1.18, 1.78, 1.78, 1.78])
        last = table.rows[-1]
        # D2：尾行 = 合计行（前 4 列并成一格，文本 `{市}合计`），表头 + 1 数据行 + 合计行
        self.assertEqual(len(table.rows), 3)
        self.assertEqual(last.cells[0].text.strip(), "测试市合计")
        self.assertIs(last.cells[0]._tc, last.cells[3]._tc)
        self.assertEqual(last.cells[4].text.strip(), table.rows[1].cells[4].text.strip())


class T2cAdviceScopeTests(unittest.TestCase):
    """T2c：工作建议结构（C1 养护提升建议/C2 迎国评前缀/C3 重点路段处治建议）与管养单位口径（C4）。"""

    @staticmethod
    def _bundle() -> dict:
        return {
            "city": "测试市",
            "marking": [],
            "height": [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K2+000",
                        "guardrail_type": "二波", "height": 600.0, "city": "测试市",
                        "manager": "明细行简称", "station_m": 1000.0, "end_m": 1001.0}],
            "bolt": [],
            "notes": [],
            "comparison_detail": [],
            "weak_segments": [{"route": "S14", "direction": "上行", "segment": "S14上行K1+000～K2+000",
                               "type": "护栏高度偏差超10cm", "reason": "三波护栏偏差超10 cm点数1个"}],
            "route_segments": [
                {"city": "测试", "route": "S14", "direction": "上行", "category": "高速公路",
                 "start": 1.0, "end": 2.0, "length": 1.0, "manager": "广东省高速公路有限公司"},
                {"city": "测试", "route": "S357", "direction": "上行", "category": "普通国省道",
                 "start": 1.0, "end": 2.0, "length": 1.0, "manager": "测试市公路事务中心"},
            ],
        }

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    @staticmethod
    def _font(paragraph):
        run = next(run for run in paragraph.runs if run.text.strip())
        rpr = run._element.find(qn("w:rPr"))
        rfonts = rpr.find(qn("w:rFonts")) if rpr is not None else None
        return ((rfonts.get(qn("w:eastAsia")) if rfonts is not None else None),
                run.font.size.pt if run.font.size else None, run.bold)

    def test_maintenance_advice_matches_template(self) -> None:
        document = self._write()
        by_text = {p.text.strip(): p for p in document.paragraphs if p.text.strip()}
        leads = ("（1）建立差异化、精准化的养护策略", "（2）分类型完善设施养护策略",
                 "（3）资金保障与技术能力提升", "（4）病害源头治理")
        titles = ("①高速公路，加强属地养护监管", "②普通国省道，统筹养护",
                  "①防护护栏", "②交通标线", "③交通标志与防眩设施")

        # C1：4 个 `（n）` 引导段 + 5 个 `①②③` 二级标题（模板-1 原文）
        for text in leads:
            self.assertIn(text, by_text, text)
            self.assertEqual(self._font(by_text[text]), ("仿宋_GB2312", 12.0, None), text)
        for text in titles:
            self.assertIn(text, by_text, text)
            self.assertEqual(by_text[text].style.name, "Heading 5", text)
            self.assertEqual(self._font(by_text[text]), ("黑体", 12.0, True), text)
        self.assertIn(engine.GD_MAINTENANCE_LEAD, by_text)

    def test_induction_prefix_and_key_route_structure(self) -> None:
        document = self._write()
        texts = [p.text.strip() for p in document.paragraphs]

        # C2：迎国评第 1 项按模板原文含 `（1）`（无句号）
        self.assertIn("（1）分类开展显性问题突击整治", texts)
        self.assertNotIn("分类开展显性问题突击整治。", texts)
        # F1：`（1）` 下 3 段按模板-1 原文逐字
        for text in ("清除标志遮挡：标志板遮挡是高速公路和普通国省道共有的突出问题。组织一次全线排查，集中修剪遮挡交通标志的树木和植被，成本低、见效快。",
                     "高速公路方面：更换变形护栏板并同步排查补齐护栏缺失螺栓，修复锈蚀、缺损防眩网；对路段旧标线未清除问题开展专项清理。",
                     "普通国省道方面：优先处置波形梁护栏缺损并补齐护栏缺失螺栓、防护设施防护能力不足隐患，快速翻新缺损严重的标线。"):
            self.assertIn(text, texts)
        body = "\n".join(texts)
        self.assertIn("（2）做好迎检路段现场排查", body)
        self.assertIn("（3）统筹力量，差异化投入", body)
        # C3：重点路段处治建议 = 模板原文单段 + 工作簿指向句，无子标题
        self.assertIn(engine.GD_KEY_ROUTE_ADVICE, texts)
        self.assertIn("优先处治路段明细见《测试市交安设施统计图表.xlsx》「优先处治路段」工作表。", texts)
        self.assertNotIn("（一）优先处治路段（6个月完成）", texts)
        self.assertNotIn("（二）闭环督办管理要求", texts)

    def test_overall_table_manager_count_matches_intro(self) -> None:
        document = self._write()
        overall = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "管理单位")
        body = "\n".join(p.text for p in document.paragraphs)
        match = re.search(r"分属(\d+)家管养单位", body)

        self.assertIsNotNone(match, body[:300])
        # C4：引言「分属 N 家管养单位」== 清单表里不同的管养单位数。
        # P1b：① 表已改为模板四行（全省/本市/非集团/集团），不再按管养单位逐行展开，
        # 所以这个数改从清单表核对 —— 两处口径同源，见 gd_manager_lookup。
        managers = {row.get("manager") for row in self._bundle()["route_segments"]
                    if row.get("category") == "高速公路"}
        self.assertEqual(int(match.group(1)), len(managers))
        # C4：② 表管养单位取路线表口径（明细行写法「明细行简称」被覆写）
        table2 = next(t for t in document.tables
                      if t.rows[0].cells[0].text.strip() == "路线编号" and "管养单位" in t.rows[0].cells[2].text)
        self.assertEqual(table2.rows[1].cells[2].text.strip(), "广东省高速公路有限公司")


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
                "manualBeforePath": "",
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
                "manualBeforePath": "",
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

    def test_run_worker_drives_real_chongqing_pipeline(self) -> None:
        """跑通未打桩的 _run_chongqing 调用链：引擎签名与桥接调用不一致必须在此暴露。"""
        import openpyxl

        summary = self.root / "real-summary.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "各区县项目概况"
        sheet.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
        sheet.append([1, "万州区", "G210", "", "一级", 2264.0, 2265.0, 1.0, 1.0])
        workbook.save(summary)

        detail = self.root / "real-detail"
        detail.mkdir()
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "高度明细"
        sheet.append(["路线编号", "方向", "原始桩号", "电子修正桩号", "护栏类型", "梁板中心高度(mm)", "异常标记"])
        sheet.append(["G210", "上行", "K2264+100", "K2264+100", "两波护栏", 600, ""])
        sheet.append(["G210", "上行", "K2264+200", "K2264+200", "两波护栏", 555, ""])
        workbook.save(detail / "重庆市-万州区-交安设施现场检测-明细.xlsx")

        disease = self.root / "real-disease"
        disease.mkdir()
        for name in ("重庆市-万州区-交安设施现场检测-病害清单.xlsx",):
            workbook.save(disease / name)

        window = FakeWindow()
        self.bridge.attach_window(window)
        payload = self.cq_payload()
        payload["values"].update({
            "summaryPath": str(summary),
            "detailPath": str(detail),
            "diseasePath": str(disease),
            "tciPath": "",
        })
        with patch.object(self.bridge, "_template_paths", return_value=self.templates()):
            self.bridge._run_worker(payload)
        events = [json.loads(script.split("desktopEvents(", 1)[1][:-1]) for script in window.scripts]
        final = events[-1]
        self.assertEqual(final["status"], "complete", final.get("message"))
        self.assertEqual(final["progress"], 100)
        # 进度条与阶段只增不减
        seen = [(event["progress"], event["stage"]) for event in events if "progress" in event]
        self.assertEqual(seen, sorted(seen))
        self.assertTrue((self.output / "重庆市万州区交安设施检测报告.xlsx").is_file())
        workbook.close()

    def test_progress_mapping(self) -> None:
        """未量化的日志只把进度推进到该阶段起点，可量化阶段按 current/total 线性插值。"""
        self.assertEqual(self.bridge._progress_from_log("正在扫描资料"), (5, 1))
        self.assertEqual(self.bridge._progress_from_log("图表已生成"), (30, 2))
        self.assertEqual(self.bridge._progress_from_log("已保存文件"), (62, 3))
        # 逐区县循环：统计占区县区间前半、报告占后半，第 2/4 个区县时进度在区间中部。
        first = self.bridge._progress_from_log(engine.format_progress("生成区县统计", 1, 4, "甲县"))
        second = self.bridge._progress_from_log(engine.format_progress("生成区县报告", 2, 4, "乙县"))
        self.assertEqual(first[1], 2)
        self.assertEqual(second[1], 3)
        self.assertGreater(second[0], first[0])

    def test_progress_never_rewinds_across_log_lines(self) -> None:
        """长耗时步骤之间的杂项日志不得让总进度或阶段回退。"""
        window = FakeWindow()
        self.bridge.attach_window(window)
        messages = [
            engine.format_progress("生成区县统计", 2, 4, "乙县"),
            "区县乙县：识别到道路编号G210 3段。",
            "统计工作簿已生成：x.xlsx",
            engine.format_progress("解析明细", 1, 16, "甲.xlsx"),
        ]
        values = []
        for message in messages:
            self.bridge._on_engine_log(message)
        events = [json.loads(script.split("desktopEvents(", 1)[1][:-1]) for script in window.scripts]
        values = [(event["progress"], event["stage"]) for event in events]
        self.assertEqual(values, sorted(values))
        self.assertEqual(values[-1][1], 2)

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
            "manualBeforePath": "",
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
            self.assertIn("共检出拼接螺栓80颗，连接螺栓120颗，缺失螺栓4颗，缺失率为1.96%", text)
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

        # G1（用户第 3 轮）：① 总体情况 下不再输出统计口径段（模板①只有 1 段）
        self.assertNotIn("覆盖整公里检测段", text)
        self.assertIn("其中两波护栏合格率", text)
        self.assertIn("三波护栏合格率", text)
        self.assertNotIn("当前区段为桥梁路段，无有效检测点位。", text)
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
            paragraphs = document.paragraphs
            start = next(i for i, p in enumerate(paragraphs) if p.text == "（2）波形梁护栏中心高度情况")
            end = next(i for i, p in enumerate(paragraphs) if p.text == "（3）波形梁护栏螺栓缺失情况")
            section = paragraphs[start:end]
            height_headings = [p.text for p in section if p.style.name == "Heading 5"]
            height_text = "\n".join(p.text for p in section)

        self.assertEqual(
            height_headings,
            [
                "①总体情况",
                "②各抽检路段情况",
                "③典型状况不佳路段及原因分析",
            ],
        )
        self.assertNotIn("覆盖整公里检测段", height_text)
        self.assertNotIn("当前区段为桥梁路段，无有效检测点位。", height_text)

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
        self.assertIn("（2）波形梁护栏中心高度情况", text)
        self.assertIn("其中两波护栏", text)
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
                "consistent": True,
            },
            {
                "indicator": "height",
                "absolute_difference": 3.0,
                "relative_deviation": None,
                "within_threshold": False,
                "consistent": False,
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


class ManualConsistencyRuleTests(unittest.TestCase):
    """M1（brief-v3 §7 用户裁决）：结果一致性 = 两侧合格判定是否相同，测值偏差/阈值不再是判据。"""

    THRESHOLDS = {"marking": 7, "height": 5, "bolt": 5}

    @classmethod
    def _compare(cls, **record):
        base = {"indicator": "height", "gtype": "三波护栏", "city": "测试市", "route": "G1",
                "direction": "上行", "segment": "K1+000~K1+100"}
        base.update(record)
        detail, summary = engine.ManualAutoComparator(cls.THRESHOLDS).compare([base])
        return detail[0], summary

    def test_same_verdict_is_consistent_for_qualified_and_failed(self):
        """两边合格同判一致、两边不合格同判一致。"""
        passed, _ = self._compare(manual=704.0, automatic=708.0)          # 697±20 内
        failed, _ = self._compare(manual=640.0, automatic=600.0)          # 697±20 外
        self.assertIs(passed["consistent"], True)
        self.assertIs(failed["consistent"], True)
        # 同一侧数值差很大仍同判（600 与 610 双波都合格）
        wide, _ = self._compare(manual=585.0, automatic=610.0, gtype="双波护栏")
        self.assertIs(wide["consistent"], True)

    def test_mixed_verdict_is_inconsistent_even_when_values_are_close(self):
        """一合一不合 → 不一致（579/582 双波：差 3mm 但跨合格界限）。"""
        detail, _ = self._compare(manual=579.0, automatic=582.0, gtype="双波护栏")
        self.assertIs(detail["within_threshold"], True)                   # 旧偏差法会判「一致」
        self.assertIs(detail["consistent"], False)                        # 新口径判「不一致」

    def test_bolt_consistency_from_missing_count_difference(self):
        """M2（第二轮裁决）：螺栓按两侧缺失数量差值判定，差值<阈值→一致、≥阈值→不一致。"""
        same, _ = self._compare(indicator="bolt", manual=3, automatic=3)          # 差值 0
        near, _ = self._compare(indicator="bolt", manual=0, automatic=4)          # 差值 4 < 5
        both_positive, _ = self._compare(indicator="bolt", manual=1, automatic=38)
        at_limit, _ = self._compare(indicator="bolt", manual=0, automatic=5)      # 差值 5 = 阈值
        far, _ = self._compare(indicator="bolt", manual=0, automatic=13)          # 差值 13
        self.assertIs(same["consistent"], True)
        self.assertIs(near["consistent"], True)
        self.assertIs(both_positive["consistent"], False)                         # 差值 37
        self.assertIs(at_limit["consistent"], False)                              # 差值 5 ≥ 5
        self.assertIs(far["consistent"], False)
        # within_threshold 同步改为「差值 < 阈值」，与 consistent 同源
        self.assertIs(near["within_threshold"], True)
        self.assertIs(far["within_threshold"], False)

    def test_bolt_threshold_is_configurable(self):
        """同一批数据换阈值结论随之改变：差值 1 在阈值 5 下判一致、阈值 1 下判不一致。"""
        def verdict(threshold):
            detail, _ = engine.ManualAutoComparator(
                {"marking": 7, "height": 5, "bolt": threshold}).compare(
                [{"indicator": "bolt", "gtype": "双波护栏", "city": "测试市", "route": "G1",
                  "direction": "上行", "segment": "K1+000~K1+100", "manual": 0, "automatic": 1}])
            return detail[0]["consistent"]
        self.assertIs(verdict(5), True)                                           # 差值 1 < 5
        self.assertIs(verdict(1), False)                                          # 差值 1 ≥ 1

    def test_bolt_missing_side_is_dash_not_consistent(self):
        """螺栓任一侧缺值 → consistent=None（Word 表显示「—」），不得视作一致。"""
        self.assertIsNone(engine.gd_consistent("bolt", None, 3, "", 5))
        self.assertIsNone(engine.gd_consistent("bolt", 2, None, "", 5))
        self.assertIsNone(engine.gd_consistent("bolt", None, None, "", 5))
        self.assertIsNone(engine.gd_pass("bolt", 0, ""))                          # 螺栓无合格判定

    def test_unknown_verdict_is_dash_and_excluded_from_rate(self):
        """任一侧无法判定 → 行级 consistent=None（显示 —），不计入一致性占比分母。"""
        unknown_type, summary = self._compare(manual=700.0, automatic=690.0, gtype="四波护栏")
        self.assertIsNone(unknown_type["consistent"])                      # 禁止 None==None 报一致
        self.assertEqual(summary["height"]["unjudged_count"], 1)
        self.assertIsNone(summary["height"]["consistency_rate"])
        judged, _ = self._compare(manual=704.0, automatic=640.0)
        self.assertIs(judged["consistent"], False)                         # 704 合格 / 640 不合格（697±20）
        _, mixed = self._compare(manual=704.0, automatic=708.0)
        self.assertEqual(mixed["height"]["consistent_count"], 1)
        self.assertEqual(mixed["height"]["unjudged_count"], 0)
        self.assertEqual(mixed["height"]["consistency_rate"], 1.0)

    def test_section_table_uses_row_consistent_and_drops_old_wording(self):
        """三张 Word 表读行级 consistent；旧「测值偏差不大于允许偏差」说明必须消失。"""
        detail, _ = engine.ManualAutoComparator(self.THRESHOLDS).compare([
            {"indicator": "height", "city": "测试市", "route": "G1", "gtype": "双波护栏",
             "segment": "K1+000~K1+100", "manual": 579.0, "automatic": 582.0},
            {"indicator": "height", "city": "测试市", "route": "G1", "gtype": "四波护栏",
             "segment": "K2+000~K2+100", "manual": 700.0, "automatic": 690.0},
        ])
        document = Document()
        tables = []
        engine.GuangdongChapterWriter._comparison_gd03_section(
            document, "高速公路", detail, self.THRESHOLDS,
            lambda headers, rows, title, merge=None, vmerge=None:
                tables.append((title, [list(row) for row in rows])))
        body = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertNotIn("测值偏差不大于允许偏差", body)
        self.assertIn("两侧的合格判定结果", body)
        # M2：旧「螺栓缺失数量合计为0判合格」口径必须消失，改为差值口径
        self.assertNotIn("合计为0判合格", body)
        self.assertIn("差值小于5记为一致", body)
        height_rows = dict(tables)["高速公路波形梁护栏中心高度人工复核对比明细表"]
        self.assertEqual([row[-1] for row in height_rows], ["不一致", "—"])
        self.assertEqual([row[3] for row in height_rows], ["600±20mm", "—"])

    def test_oscillation_marking_remark_noted_below_table(self):
        """M3（brief-v3 §7）：源表 备注「现场震荡标线」的标线行在表下说明段列出桩号区段；无备注行则不输出该句。"""
        def render(remark):
            detail, _ = engine.ManualAutoComparator(self.THRESHOLDS).compare([
                {"indicator": "marking", "city": "测试市", "route": "G4", "direction": "下行",
                 "segment": "K1887+200~K1887+100", "manual": 73.2, "automatic": 190.0, "remark": remark},
                {"indicator": "marking", "city": "测试市", "route": "G0422", "direction": "下行",
                 "segment": "K699+200~K699+100", "manual": 38.4, "automatic": 184.1, "remark": remark},
            ])
            document = Document()
            tables = []
            engine.GuangdongChapterWriter._comparison_gd03_section(
                document, "高速公路", detail, self.THRESHOLDS,
                lambda headers, rows, title, merge=None, vmerge=None:
                    tables.append((title, list(headers), [list(row) for row in rows])))
            return document, tables

        document, tables = render("现场震荡标线")
        body = "\n".join(paragraph.text for paragraph in document.paragraphs)
        self.assertIn("注：以上 2 个路段现场为震荡标线，其人工复核与自动化检测结果的差异受标线型式影响。", body)
        self.assertIn("G4下行K1887+200-K1887+100", body)          # 桩号连接符归一化为 ASCII `-`
        self.assertIn("G0422下行K699+200-K699+100", body)
        self.assertNotIn("～", body)
        # 约束：不新增表列 —— 标线表列头与模板固定 10 列一字不动
        title, headers, _ = tables[0]
        self.assertEqual(title, "高速公路标线逆反射亮度系数人工复核对比明细表")
        self.assertEqual(headers, ["路线编号", "桩号区段", "标线颜色", "人工检测复核结果", "测值",
                                   "合格判定", "自动化检测结果", "测值", "合格判定", "结果一致性"])

        document, _ = render("")
        self.assertNotIn("现场为震荡标线", "\n".join(p.text for p in document.paragraphs))

    def test_read_file_bolt_blank_side_is_missing_not_zero(self):
        """螺栓某侧两格皆空 → 该侧 None（进 issues），不得静默补 0 判合格。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "进场后_人工自动化对比_统计结果.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "螺栓缺失"
            # 与真实表一致：第1行主表头，第2行子表头，数据自第3行起
            sheet.append(["地区", "路线", "护栏类型", "护栏位置", "方向", "桩号范围",
                          "人工复核螺栓缺失数量", None, "自动化螺栓缺失数量", None, "备注"])
            sheet.append([None, None, None, None, None, None,
                          "拼接螺栓缺失数量", "连接螺栓缺失数量",
                          "拼接螺栓缺失数量", "连接螺栓缺失数量", None])
            sheet.append(["测试市", "G1", "双波护栏", "路侧", "上行", "K1+000~K1+100", None, None, 2, 0, None])
            workbook.save(path)
            records, issues = engine.ManualAutoComparator.read_file(str(path))
        self.assertEqual(records, [])
        self.assertTrue(any("字段缺失" in issue for issue in issues), issues)


class Gd03ManualReviewAdviceTests(unittest.TestCase):
    """GD03（4）人工复核对比情况与（三）工作建议：对标附件的五级标题与表列名。"""

    @staticmethod
    def _detail_records():
        return [
            {"indicator": "marking", "city": "云浮市", "route": "G2518", "direction": "上行",
             "segment": "K230+200~K230+100", "category": "高速公路", "source": "正式检测后",
             "manual": 105.76, "automatic": 101.04},
            {"indicator": "marking", "city": "云浮市", "route": "S51", "direction": "上行",
             "segment": "K59+100~K59+200", "category": "高速公路", "source": "正式检测后",
             "manual": 26.51, "automatic": 27.15},
            {"indicator": "height", "city": "云浮市", "route": "G2518", "direction": "上行",
             "segment": "K230+200~K230+100", "category": "高速公路", "source": "正式检测后",
             "gtype": "三波护栏", "manual": 704.12, "automatic": 708.52},
            {"indicator": "height", "city": "云浮市", "route": "S51", "direction": "上行",
             "segment": "K59+200~K59+100", "category": "高速公路", "source": "正式检测后",
             "gtype": "三波护栏", "manual": 705.90, "automatic": 700.28},
            {"indicator": "bolt", "city": "云浮市", "route": "G2518", "direction": "上行",
             "segment": "K231~K230", "category": "高速公路", "source": "正式检测后",
             "gtype": "三波护栏", "manual": 1, "automatic": 38,
             "msplice": 1, "mconn": 0, "asplice": 30, "aconn": 8},
        ]

    def _bundle(self):
        detail, summary = engine.ManualAutoComparator(
            {"marking": 7, "height": 5, "bolt": 5}).compare(self._detail_records())
        return {
            "city": "云浮市",
            "marking": [
                {"city": "云浮市", "route": "G2518", "direction": "下行", "station_m": 343000,
                 "manager": "广东省路桥建设发展有限公司云梧分公司", "category": "高速公路"},
                {"city": "云浮市", "route": "G324", "direction": "上行", "station_m": 1182000,
                 "manager": "云浮市云城区公路事务中心", "category": "普通国省道"},
            ],
            "height": [], "bolt": [],
            "weak_segments": [
                {"route": "G2518", "direction": "下行", "segment": "G2518下行K0343～K0308",
                 "type": "护栏高度偏差超10cm", "reason": "三波护栏偏差超10 cm点数120个"},
                {"route": "G324", "direction": "上行", "segment": "K1182+000~K1186+880",
                 "type": "标线连续3km不合格", "marking_position": "左侧标线",
                 "start_m": 1182000, "end_m": 1186880, "reason": "连续不合格长度达4.88 km"},
            ],
            "comparison_detail": detail,
            "comparison_summary": summary,
        }

    @staticmethod
    def _collector():
        tables = []

        def table(headers, rows, title, merge=None, vmerge=None):
            tables.append({"title": title, "headers": list(headers), "rows": [list(row) for row in rows],
                           "merge": merge, "vmerge": vmerge})

        return tables, table

    def test_comparison_section_matches_attachment_structure(self):
        bundle = self._bundle()
        document = Document()
        tables, table = self._collector()
        engine.GuangdongChapterWriter._comparison_gd03_section(
            document, "高速公路", bundle["comparison_detail"], {"marking": 7, "height": 5, "bolt": 5}, table)
        headings = [p.text for p in document.paragraphs if p.style.name == "Heading 5"]
        self.assertEqual(headings, [])
        body = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("对比分析云浮市高速公路抽检路段标线逆反射亮度系数（2个路段）、波形梁护栏中心高度（2个路段）"
                      "和螺栓缺失情况（1个路段）", body)
        self.assertIn("各指标明细对比情况详见下表。", body)
        self.assertIn("区间，自动化检测结果的病害数量多于人工复核结果的原因，"
                      "是人工复核前管养单位已针对该区间开展了病害整改作业，可处置的缺陷均已完成补装螺栓处理。", body)
        for token in ("1.1 偏差指标", "1.2 判定标准", "1.3 一致性口径", "2.1 偏差范围",
                      "3.2 一致性", "4.2 检出一致性", "小结：", "偏差范围表"):
            self.assertNotIn(token, body)
        titles = [item["title"] for item in tables]
        self.assertEqual(titles, [
            "高速公路标线逆反射亮度系数人工复核对比明细表",
            "高速公路波形梁护栏中心高度人工复核对比明细表",
            "高速公路路侧波形梁护栏螺栓缺失人工复核对比明细表"])
        by_title = {item["title"]: item for item in tables}
        # D14：三表均为两行表头（首行合并），列头逐字照模板
        self.assertEqual(by_title["高速公路标线逆反射亮度系数人工复核对比明细表"]["headers"],
                         ["路线编号", "桩号区段", "标线颜色", "人工检测复核结果", "测值", "合格判定",
                          "自动化检测结果", "测值", "合格判定", "结果一致性"])
        self.assertEqual(by_title["高速公路波形梁护栏中心高度人工复核对比明细表"]["headers"],
                         ["路线编号", "桩号区段", "护栏类型", "合格值(mm)", "人工检测复核结果", "测值", "合格判定",
                          "自动化检测结果", "测值", "合格判定", "结果一致性"])
        self.assertEqual(by_title["高速公路路侧波形梁护栏螺栓缺失人工复核对比明细表"]["headers"],
                         ["路线", "桩号区段", "护栏类型", "人工检测复核结果", "拼接螺栓缺失", "连接螺栓缺失", "总体缺失",
                          "自动化检测结果", "拼接螺栓缺失", "连接螺栓缺失", "总体缺失", "结果一致性"])
        self.assertEqual(by_title["高速公路标线逆反射亮度系数人工复核对比明细表"]["merge"],
                         {3: ("人工检测复核结果", 3), 6: ("自动化检测结果", 3)})
        marking_row = by_title["高速公路标线逆反射亮度系数人工复核对比明细表"]["rows"][0]
        self.assertEqual(marking_row[0], "G2518")
        self.assertEqual(marking_row[3:], ["105.76", "合格", "101.04", "合格", "一致"])
        bolt_row = by_title["高速公路路侧波形梁护栏螺栓缺失人工复核对比明细表"]["rows"][0]
        # M2：螺栓无合格判定列，按两侧缺失数量差值（|1−38|=37 ≥ 阈值 5）判不一致
        self.assertEqual(bolt_row[3:], ["1", "0", "1", "30", "8", "38", "不一致"])

    def test_inspection_segments_aggregate_by_category_route_direction(self):
        bundle = {"marking": [{"category": "高速公路", "route": "G4", "direction": "上行", "station_m": 100.0, "manager": "甲分公司"},
                              {"category": "高速公路", "route": "G4", "direction": "上行", "station_m": 5100.0}],
                  "bolt": [{"category": "普通国省道", "route": "S6", "direction": "下行", "station_m": 200.0, "manager": "乙中心"}]}
        rows = engine.gd_inspection_segments(bundle)
        self.assertEqual([(row["category"], row["route"], row["direction"]) for row in rows],
                         [("高速公路", "G4", "上行"), ("普通国省道", "S6", "下行")])
        self.assertEqual((rows[0]["start_text"], rows[0]["end_text"], rows[0]["length_text"]), ("0.1", "5.1", "5"))
        self.assertEqual(rows[0]["manager"], "甲分公司")

    def test_inspection_segments_prefer_route_table(self):
        bundle = {"city": "东莞市", "marking": [{"category": "高速公路", "route": "G4", "direction": "下行", "station_m": 1000.0}],
                  "route_segments": [{"city": "东莞", "route": "G220", "direction": "上行", "category": "普通国省道",
                                      "start": 2528.0, "end": 2530.0, "length": 2.0, "manager": "东莞市公路事务中心"}]}
        rows = engine.gd_inspection_segments(bundle)
        self.assertEqual(len(rows), 1)
        self.assertEqual((rows[0]["route"], rows[0]["start_text"], rows[0]["length_text"]), ("G220", "2528", "2"))

    def test_inspection_segments_review_remark_from_phases(self):
        bundle = {"city": "东莞市",
                  "review_phases": {"进场前": {("G220", "上行"): [(2558000, 2564000)]}},
                  "route_segments": [{"city": "东莞", "route": "G220", "direction": "上行", "category": "普通国省道",
                                      "start": 2558.0, "end": 2564.0, "length": 6.0, "manager": "东莞市公路事务中心"}]}
        rows = engine.gd_inspection_segments(bundle)
        self.assertEqual(rows[0]["remark"], "人工复核：进场前6公里")

    def test_inspection_segments_skip_rows_without_category(self):
        """P1g：类别判不出（None/空）的明细行不进抽检路段清单，避免出现空 `类型` 行。"""
        bundle = {"marking": [{"category": "高速公路", "route": "S2", "direction": "上行", "station_m": 100.0},
                              {"category": None, "route": "S272", "direction": "上行", "station_m": 29000.0},
                              {"category": "  ", "route": "S278", "direction": "上行", "station_m": 30000.0}],
                  "height": [{"category": "", "route": "S282", "direction": "下行", "station_m": 31000.0}]}
        rows = engine.gd_inspection_segments(bundle)
        self.assertEqual([(row["category"], row["route"]) for row in rows], [("高速公路", "S2")])

    def test_advice_section_tables_and_headings(self):
        bundle = self._bundle()
        document = Document()
        tables, table = self._collector()
        engine.GuangdongChapterWriter._advice_gd03_section(document, bundle, table)
        headings = [p.text for p in document.paragraphs if p.style.name == "Heading 3"]
        self.assertEqual(headings, ["1.重点路段处治建议（如有）", "2.迎国评工作建议", "3.养护提升建议"])
        # D16/Q1：模板无优先处治路段表，正文改为清单说明句（明细进交安设施统计图表工作簿）
        self.assertEqual(tables, [])
        body = "\n".join(p.text for p in document.paragraphs)
        self.assertIn("优先处治路段明细见《云浮市交安设施统计图表.xlsx》「优先处治路段」工作表。", body)
        for token in ("清除标志遮挡", "高速公路方面", "普通国省道方面",
                      "（2）做好迎检路段现场排查", "（3）统筹力量，差异化投入"):
            self.assertIn(token, body)
        # C3：模板该节仅 1 个正文段，原（一）/（二）子标题与督办段落已删（明细进工作簿）
        for removed in ("（一）优先处治路段（6个月完成）", "（二）闭环督办管理要求", "建立整改清单", "验收销号", "考核挂钩"):
            self.assertNotIn(removed, body)


class GuangdongRouteTableFormatTests(unittest.TestCase):
    """路线分类表按表头识别多种格式（省检线路统计 / 附件3_4 / 旧格式），不依赖文件名。
    关键口径：块标记列（高速/国省道）与道路等级（一级/二级）都要能判类别，汇总行不计入。"""

    @staticmethod
    def _write(path: Path, sheets: dict) -> Path:
        from openpyxl import Workbook
        wb = Workbook()
        wb.remove(wb.active)
        for name, rows in sheets.items():
            ws = wb.create_sheet(name)
            for row in rows:
                ws.append(list(row))
        wb.save(path)
        return path

    def test_block_marker_column_and_level_column_and_headerless_sheet(self):
        """三种工作表形态同时出现在一个文件里，均应被正确解析。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write(Path(temp_dir) / "routes.xlsx", {
                # 有表头 + 无表头首列块标记（高速/国省道），地市逐行填写
                "珠三角片区": [
                    ["", "序号", "地市", "路线", "方向", "路段起点", "路段终点", "检测里程（km）", "管养单位"],
                    ["高速", 1, "广州", "G0421", "下行", 1305, 1365, 60, "A公司"],
                    ["国省道", 2, "广州", "G105", "上行", 2482, 2483, 1, "B中心"],
                    ["国省道", 3, "佛山", "S123", "上行", 28, 29, 1, "C站"],
                    ["总计", "", "", "", "", "", "", 62, ""],
                ],
                # 有表头 + “线路号”别名 + 道路等级列 + 地市向下填充
                "西片区": [
                    ["", "地市", "线路号", "线路名", "方向", "起点桩号", "终点桩号", "里程", "道路等级", "管养单位"],
                    ["", "韶关市", "G0422", "武汉－深圳", "下行", 650, 700, 50, "高速公路", "D处"],
                    ["", "", "G323", "瑞金－清水河", "上行", 288, 293, 5, "一级", "E中心"],
                    ["", "清远市", "S292", "白土-石角", "上行", 79, 91, 12, "二级", "F中心"],
                    ["", "总计", "", "", "", "", "", 67, "", ""],
                ],
                # 无表头：固定列序，首列地市向下填充
                "Sheet3": [
                    ["韶关", "G105", "北京－澳门", "上行", 2391, 2399, "", "一级", "G中心"],
                    ["", "G323", "瑞金－清水河", "上行", 288, 293, "", "二级", "H中心"],
                    ["肇庆", "G321", "广州－成都", "上行", 101, 106, "", "一级", "I中心"],
                ],
            })
            index = engine.RouteCategoryIndex.from_file(path)

        self.assertEqual(index.category("广州", "G0421"), "高速公路")
        self.assertEqual(index.category("广州", "G105"), "普通国省道")
        self.assertEqual(index.category("佛山", "S123"), "普通国省道")
        self.assertEqual(index.category("韶关", "G0422"), "高速公路")
        self.assertEqual(index.category("韶关", "G323"), "普通国省道")
        self.assertEqual(index.category("清远", "S292"), "普通国省道")
        self.assertEqual(index.category("韶关", "G105"), "普通国省道")
        self.assertEqual(index.category("肇庆", "G321"), "普通国省道")
        # “韶关市/韶关”两种写法归一后指向同一地市
        self.assertEqual(engine.RouteCategoryIndex._norm_city("韶关市"), "韶关")
        # 汇总行不得进入分类
        with self.assertRaises(KeyError):
            index.category("总计", "G0421")

    def test_legacy_and_annex_formats_still_parse(self):
        """旧格式（道路类别列）与附件3_4 格式（工作表名区分类别）不得回归。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            legacy = self._write(Path(temp_dir) / "legacy.xlsx", {
                "任意表": [
                    ["地市", "路线", "道路类别"],
                    ["云浮市", "G2518", "高速公路"],
                    ["云浮市", "G324", "普通国省道"],
                ],
            })
            annex = self._write(Path(temp_dir) / "annex.xlsx", {
                "附件3-高速公路明细": [
                    ["2026年省检高速公路明细"],
                    ["地市", "路线编码", "起点桩号", "终点桩号"],
                    ["云浮市", "G80", 100, 120],
                ],
                "附件4-普通国省道明细": [
                    ["2026年省检普通国省道明细"],
                    ["地市", "路线编码", "起点桩号", "终点桩号"],
                    ["云浮市", "G324", 1300, 1320],
                ],
            })
            legacy_index = engine.RouteCategoryIndex.from_file(legacy)
            annex_index = engine.RouteCategoryIndex.from_file(annex)

        self.assertEqual(legacy_index.category("云浮市", "G2518"), "高速公路")
        self.assertEqual(legacy_index.category("云浮", "G324"), "普通国省道")
        self.assertEqual(annex_index.category("云浮", "G80"), "高速公路")
        self.assertEqual(annex_index.category("云浮", "G324"), "普通国省道")

    def test_owner_column_is_read_and_exported_as_fifth_column(self):
        """P1d：表头含「经营主体」时读入为第 5 字段，owner_for 按 (归一地市, 路线号) 取值。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write(Path(temp_dir) / "routes.xlsx", {
                "路线分类表": [
                    ["地市", "路线编号", "路线名称", "道路类别", "经营主体"],
                    ["东莞", "G0422", "武汉-深圳", "高速公路", "集团"],
                    ["珠海", "S270", "金湾高速", "普通国省道", None],
                ],
            })
            index = engine.RouteCategoryIndex.from_file(path)
            # 缺列不抛异常；地市去后缀归一后仍能取到（东莞/东莞市、珠海/珠海市）
            self.assertEqual(index.owner_for("东莞", "G0422"), "集团")
            self.assertEqual(index.owner_for("东莞市", "g0422"), "集团")
            self.assertEqual(index.owner_for("珠海", "S270"), None)
            self.assertEqual(index.owner_for("韶关", "G105"), None)

            out = Path(temp_dir) / "out"
            written = engine.write_guangdong_route_workbook_all(
                index, {"G0422": "武汉-深圳", "S270": "金湾高速"}, out, filename="含经营主体.xlsx")
            self.assertIsNotNone(written)
            from openpyxl import load_workbook
            sheet = load_workbook(written, data_only=True)["路线分类表"]
            rows = list(sheet.iter_rows(values_only=True))
        self.assertEqual(rows[0], ("地市", "路线编号", "路线名称", "道路类别", "经营主体"))
        self.assertEqual(len(rows[0]), 5)
        by_route = {r[1]: r for r in rows[1:]}
        self.assertEqual(by_route["G0422"][4], "集团")
        # 取不到 → 空单元格，不写「—」「缺失」等占位
        self.assertIsNone(by_route["S270"][4])
        # 前 4 列顺序与语义不变
        self.assertEqual(by_route["G0422"][:4], ("东莞", "G0422", "武汉-深圳", "高速公路"))

    def test_owner_falls_back_to_annex_match_when_column_missing(self):
        """P1d：表里没有经营主体列时回退附件匹配；跨主体且无桩号 → None（不得按路线号猜）。"""
        index = engine.RouteCategoryIndex({("东莞", "G0422"): "高速公路", ("东莞", "S6"): "高速公路"})
        self.assertEqual(index.owners, {})
        index.owner_index = {
            ("东莞", "G0422"): [{"start": 0.0, "end": 10.0, "owner": "集团", "keeper": "A"}],
            # 同一路线跨主体：有桩号才能消歧
            ("东莞", "S6"): [{"start": 0.0, "end": 13.459, "owner": "集团", "keeper": "A"},
                             {"start": 13.459, "end": 77.869, "owner": "非集团", "keeper": "B"}],
        }
        self.assertEqual(index.owner_for("东莞市", "G0422"), "集团")
        self.assertIsNone(index.owner_for("东莞", "S6"))
        self.assertEqual(index.owner_for("东莞", "S6", start=20, end=30), "非集团")
        self.assertEqual(index.owner_for("东莞", "S6", start=1, end=5), "集团")
        # 附件无此(市,路线) → None
        self.assertIsNone(index.owner_for("东莞", "G105"))

    def test_city_normalization_strips_trailing_shi_for_folder_derived_names(self):
        """P1d：现有输入表混用「韶关市」与「中山」，按文件夹推导的市名无后缀，归一后必须指向同一键。"""
        self.assertEqual(engine.RouteCategoryIndex._norm_city("韶关市"), "韶关")
        self.assertEqual(engine.RouteCategoryIndex._norm_city(" 中山 "), "中山")
        self.assertEqual(engine.RouteCategoryIndex._norm_city("江门市"), "江门")
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write(Path(temp_dir) / "routes.xlsx", {
                "路线分类表": [
                    ["地市", "路线编号", "路线名称", "道路类别", "经营主体"],
                    ["韶关市", "G105", "北京－澳门", "普通国省道", None],
                    ["中山", "G105", "北京－澳门", "普通国省道", None],
                ],
            })
            index = engine.RouteCategoryIndex.from_file(path)
        # 两种写法都能按归一后的键查到，且是同一条记录
        self.assertEqual(index.category("韶关", "G105"), "普通国省道")
        self.assertEqual(index.category("韶关市", "G105"), "普通国省道")
        self.assertEqual(index.rows("韶关"), index.rows("韶关市"))
        self.assertEqual([row["route"] for row in index.rows("中山")], ["G105"])

    # ---------- P1d-b ----------

    def test_category_for_records_refuses_to_borrow_other_city_route(self):
        """P1d-b P0-A：第 2 级兜底不得跨市借类别（江门明细里的 12 个他市省道码曾污染江门①②分母）。"""
        index = engine.RouteCategoryIndex({
            ("江门", "S269"): "普通国省道", ("江门", "G15"): "高速公路",
            ("佛山", "S272"): "普通国省道", ("云浮", "S368"): "普通国省道",
        })
        # 他市路线码：全表唯一命中，但归属地市 ≠ 当前市 → 拒绝借用，返回 None（与留空行走同一出口）
        self.assertIsNone(index.category_for_records("江门", "S272"))
        self.assertIsNone(index.category_for_records("江门", "S368"))
        # 本市路线号照常命中（贯通路线场景不受影响）
        self.assertEqual(index.category_for_records("江门", "S269"), "普通国省道")
        self.assertEqual(index.category_for_records("江门市", "s269"), "普通国省道")
        self.assertEqual(index.category_for_records("江门", "G15"), "高速公路")
        # 路线号在多市出现且类别唯一时，只有当前市自己那行才算命中
        index2 = engine.RouteCategoryIndex({("佛山", "G105"): "普通国省道", ("中山", "G105"): "普通国省道"})
        self.assertEqual(index2.category_for_records("中山", "G105"), "普通国省道")
        self.assertIsNone(index2.category_for_records("东莞", "G105"))

    def test_category_for_records_falls_back_to_route_name(self):
        """P1d-b P0-B：源表把「路线编号」列整列填成路线名称时，按本市唯一的名称反查补回类别。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write(Path(temp_dir) / "routes.xlsx", {
                "路线分类表": [
                    ["地市", "路线编号", "路线名称", "道路类别"],
                    ["广州", "S2", "广河高速", "高速公路"],
                    ["惠州", "S2", "广河高速", "高速公路"],
                    ["广州", "G4", "北京－港澳", "高速公路"],
                ],
            })
            index = engine.RouteCategoryIndex.from_file(path)
        # 广州「广河高速」在本市唯一对应 S2 → 补回高速公路（此前整段 32256 行被排除出分母）
        self.assertEqual(index.category_for_records("广州", "广河高速"), "高速公路")
        self.assertEqual(index.category_for_records("广州市", " 广河高速 "), "高速公路")
        self.assertEqual(index.category_for_records("惠州", "广河高速"), "高速公路")
        # 名称在两市都出现 → 按 (地市, 名称) 各自唯一命中，不串市
        self.assertEqual(index.category_for_records("东莞", "广河高速"), None)
        # 本市同名多行 → 不做名称兜底
        with tempfile.TemporaryDirectory() as temp_dir:
            dup = self._write(Path(temp_dir) / "dup.xlsx", {
                "路线分类表": [
                    ["地市", "路线编号", "路线名称", "道路类别"],
                    ["广州", "S2", "广河高速", "高速公路"],
                    ["广州", "S81", "广河高速", "高速公路"],
                ],
            })
            dup_index = engine.RouteCategoryIndex.from_file(dup)
        self.assertNotIn(("广州", "广河高速"), dup_index.route_of_name)
        self.assertEqual(dup_index.category_for_records("广州", "广河高速"), None)

    def test_route_code_errata_only_rewrites_the_declared_pair(self):
        """P1d-b：茂名/G291 是录入笔误，实为 S291（用户 2026-09-27 澄清）；其它市/其它码不受影响。"""
        self.assertEqual(engine.ROUTE_CODE_ERRATA, {("茂名", "G291"): "S291"})
        index = engine.RouteCategoryIndex({
            ("茂名", "S291"): "普通国省道", ("茂名", "S292"): "普通国省道",
            ("佛山", "G291"): "普通国省道",
        })
        # 归一到 S291 → 命中分类表既有行，不得落成类别留空
        self.assertEqual(index.category_for_records("茂名", "G291"), "普通国省道")
        self.assertEqual(index.category_for_records("茂名市", " g291 "), "普通国省道")
        self.assertEqual(index.category_or_none("茂名", "G291"), "普通国省道")
        # 其它市不误伤
        self.assertEqual(index.category_for_records("佛山", "G291"), "普通国省道")
        # 茂名的其它码不受影响
        self.assertEqual(index.category_for_records("茂名", "S292"), "普通国省道")
        self.assertIsNone(index.category_for_records("茂名", "G292"))

    def test_city_display_appends_shi_for_render_layer_only(self):
        """P1d-b 裁决①：显示层市名统一带「市」；分类表「地市」列与 _norm_city 匹配键保持原样。"""
        self.assertEqual(engine.RouteCategoryIndex.city_display("东莞"), "东莞市")
        self.assertEqual(engine.RouteCategoryIndex.city_display("云浮市"), "云浮市")
        self.assertEqual(engine.RouteCategoryIndex.city_display(" 韶关市 "), "韶关市")
        self.assertEqual(engine.RouteCategoryIndex.city_display(""), "")
        # 匹配键不受影响
        self.assertEqual(engine.RouteCategoryIndex._norm_city("东莞市"), "东莞")
        with tempfile.TemporaryDirectory() as temp_dir:
            path = self._write(Path(temp_dir) / "routes.xlsx", {
                "路线分类表": [
                    ["地市", "路线编号", "路线名称", "道路类别"],
                    ["东莞", "S6", "广深沿江", "高速公路"],
                    ["江门市", "G15", "沈阳－海口", "高速公路"],
                ],
            })
            index = engine.RouteCategoryIndex.from_file(path)
            out = Path(temp_dir) / "out"
            engine.write_guangdong_route_workbook_all(index, {}, out, filename="r.xlsx")
            from openpyxl import load_workbook
            rows = list(load_workbook(out / "r.xlsx", data_only=True)["路线分类表"].iter_rows(values_only=True))
        # 表内「地市」列原样保留（东莞 / 江门市），不因显示层统一而改写
        self.assertEqual(sorted(r[0] for r in rows[1:]), ["东莞", "江门市"])
        # 表里没有阳江 → 不补 EXTRA_ROUTE_ROWS（部分表/单市表不凭空多行）
        self.assertEqual(len(rows[1:]), 2)

    def test_export_adds_missing_extra_row_but_never_duplicates_it(self):
        """P1d-b：阳江/G324 只在分类表缺该行时补一次；本工具产物回灌时不重复补。"""
        header = ["地市", "路线编号", "路线名称", "道路类别"]
        with tempfile.TemporaryDirectory() as temp_dir:
            missing_path = self._write(Path(temp_dir) / "a.xlsx", {"路线分类表": [
                header + [], ["阳江市", "G234", "兴隆－阳江", "普通国省道"],
                ["阳江市", "S278", "罗镜-溪头", "普通国省道"],
            ]})
            first = engine.RouteCategoryIndex.from_file(missing_path)
            out = Path(temp_dir) / "o1"
            engine.write_guangdong_route_workbook_all(first, {}, out, filename="r.xlsx")
            from openpyxl import load_workbook
            rows1 = list(load_workbook(out / "r.xlsx", data_only=True)["路线分类表"].iter_rows(values_only=True))
            # 补一行 → 阳江 3 行（含 G324）
            self.assertEqual(len(rows1[1:]), 3)
            self.assertEqual(sorted((r[0], r[1]) for r in rows1[1:]),
                             [("阳江市", "G234"), ("阳江市", "G324"), ("阳江市", "S278")])
            # 拿产物回灌 → 不得再补（不得出现重复行）
            second = engine.RouteCategoryIndex.from_file(out / "r.xlsx")
            out2 = Path(temp_dir) / "o2"
            engine.write_guangdong_route_workbook_all(second, {}, out2, filename="r.xlsx")
            rows2 = list(load_workbook(out2 / "r.xlsx", data_only=True)["路线分类表"].iter_rows(values_only=True))
        self.assertEqual(len(rows2[1:]), 3)
        self.assertEqual(len({(r[0], r[1]) for r in rows2[1:]}), 3)

    def test_export_fills_owner_by_measured_segments_and_marks_cross_owner(self):
        """P1d-b 裁决③：分类表无桩号列 → 用明细文件名实测段消歧；跨两个主体写「跨主体」。"""
        index = engine.RouteCategoryIndex({("东莞", "S6"): "高速公路", ("东莞", "G4"): "高速公路",
                                           ("东莞", "S269"): "普通国省道"})
        index.owner_index = {
            ("东莞", "S6"): [{"start": 0.0, "end": 13.459, "owner": "集团", "keeper": "A"},
                             {"start": 13.459, "end": 77.869, "owner": "非集团", "keeper": "B"}],
            ("东莞", "G4"): [{"start": 0.0, "end": 20.0, "owner": "集团", "keeper": "A"},
                             {"start": 20.0, "end": 40.0, "owner": "非集团", "keeper": "B"}],
        }
        # 无实测段 → 跨主体判不出，仍留空
        self.assertIsNone(index.owner_for("东莞", "S6"))
        # 实测段全部落入同一主体 → 写该主体
        index.measured = {("东莞", "S6"): [(20.0, 30.0)], ("东莞", "G4"): [(0.0, 18.0), (20.0, 40.0)]}
        self.assertEqual(index.owner_for("东莞", "S6"), "非集团")
        self.assertEqual(index.owner_for("东莞", "G4"), engine.CROSS_OWNER)
        # 普通国省道一律留空，即使附件里有分段也不写
        index.owner_index[("东莞", "S269")] = [{"start": 0.0, "end": 5.0, "owner": "集团", "keeper": "A"}]
        self.assertEqual(index.owner_for("东莞", "S269"), "集团")
        with tempfile.TemporaryDirectory() as temp_dir:
            out = Path(temp_dir) / "out"
            engine.write_guangdong_route_workbook_all(index, {}, out, filename="r.xlsx")
            from openpyxl import load_workbook
            rows = list(load_workbook(out / "r.xlsx", data_only=True)["路线分类表"].iter_rows(values_only=True))
        by_route = {r[1]: r for r in rows[1:]}
        self.assertEqual(by_route["S6"][4], "非集团")
        self.assertEqual(by_route["G4"][4], "跨主体")
        self.assertIsNone(by_route["S269"][4])      # 普通国省道：空单元格，不写占位

    def test_cross_owner_row_splits_by_station_and_closes_the_gap(self):
        """P1h：跨主体行按 P2 桩号拆分，①表「非集团+集团=全市」闭合（广州 433+75=508≠655 的根因修复）。

        数据取自广州 G0425：附件在 K18.339 处换主体，抽检行是 [0, 40]。
        拆分前该行整行落空（两侧都不计）→ 差额 40km；拆分后 18.339+21.661=40.000 归零。
        """
        index = engine.RouteCategoryIndex({("广州", "G0425"): "高速公路", ("广州", "S2"): "高速公路"},
                                          owners={("广州", "G0425"): engine.CROSS_OWNER},
                                          names={("广州", "S2"): "广河高速"})
        index.owner_index = {("广州", "G0425"): [
            {"start": 0.0, "end": 18.339, "owner": "非集团", "keeper": "广州公路"},
            {"start": 18.339, "end": 54.371, "owner": "集团", "keeper": "京珠北段"}]}
        # 行级 owner：P2 覆盖两个主体 → 判不出（不是取第一段/多数）
        self.assertIsNone(index.owner_for("广州", "G0425", 0, 40, "广州公路"))
        spans = index.owner_spans("广州", "G0425", 0, 40)
        self.assertEqual(spans, [(0.0, 18.339, "非集团"), (18.339, 40.0, "集团")])
        self.assertEqual(engine.gd_cross_owner_spans(spans, 0, 40), spans)

        row = {"route": "G0425", "direction": "下行", "category": "高速公路",
               "start": 0.0, "end": 40.0, "length_km": 40.0, "owner": None, "owner_spans": spans}
        plain = {"route": "S2", "direction": "下行", "category": "高速公路",
                 "start": 0.0, "end": 70.0, "length_km": 70.0, "owner": "非集团"}
        inspection = [row, plain]
        # 里程：拆分行各归其主体，合计回到全市 110 km，差额 0.000
        other = engine.gd_mileage(inspection, category="高速公路", owner="非集团")
        prov = engine.gd_mileage(inspection, category="高速公路", owner="集团")
        self.assertAlmostEqual(other, 18.339 + 70.0, places=6)
        self.assertAlmostEqual(prov, 21.661, places=6)
        self.assertAlmostEqual(other + prov, 110.0, places=6)      # 差额 0

        # 率分母：明细行按自身 station_m（米）落段，两侧不重不漏
        owner_map = engine.gd_owner_map(inspection)
        details = [{"route": "G0425", "direction": "下行", "station_m": 5000.0},     # K5  非集团
                   {"route": "G0425", "direction": "下行", "station_m": 30000.0},   # K30 集团
                   {"route": "G0425", "direction": "下行", "station_m": 40000.0}]   # K40 末点
        self.assertEqual(len(engine.gd_owner_rows(details, owner_map, "非集团")), 1)
        self.assertEqual(len(engine.gd_owner_rows(details, owner_map, "集团")), 2)

    def test_owner_gap_note_is_visible_in_report_text(self):
        """P1i：①表差额必须有产物内可见的说明句（闭合写事实句，有差额写未定主体的路线/里程/原因）。"""
        closed = [{"route": "G0422", "direction": "下行", "category": "高速公路",
                   "start": 0.0, "end": 33.0, "length_km": 33.0, "owner": "非集团"}]
        note = engine.gd_owner_gap_note(closed, "高速公路")
        self.assertIsNotNone(note)
        self.assertIn("经营主体已全部核定", note or "")
        self.assertIn("33.000", note or "")

        gap = [{"route": "G0422", "direction": "下行", "category": "高速公路",
                "start": 0.0, "end": 33.0, "length_km": 33.0, "owner": None}]
        note = engine.gd_owner_gap_note(gap, "高速公路")
        self.assertIsNotNone(note)
        self.assertIn("少 33.000 km", note or "")
        self.assertIn("G0422", note or "")       # 点名未核定主体的路段
        self.assertIn("未覆盖", note or "")      # 写明原因，不空话
        self.assertIsNone(engine.gd_owner_gap_note([], "高速公路"))

    def test_branch_key_matches_same_branch_written_differently(self):
        """P1i：管养单位按分支名归一 —— 同一分公司在附件/明细里公司名不同仍视为同一单位。

        深圳 G0422：附件写「广东省公路建设有限公司博深分公司」，明细写「广东博大高速公路
        有限公司博深分公司」。整串等值时 P3 落空，附件未覆盖的 7.131km 判不出主体。
        """
        self.assertEqual(engine.gd_branch_key("广东博大高速公路有限公司博深分公司"),
                         engine.gd_branch_key("广东省公路建设有限公司博深分公司"))
        # 不同分支不得混同
        self.assertNotEqual(engine.gd_branch_key("广东省公路建设有限公司博深分公司"),
                            engine.gd_branch_key("深圳高速运营发展有限公司"))
        self.assertNotEqual(engine.gd_branch_key("广东省公路建设有限公司江罗分公司"),
                            engine.gd_branch_key("广东省公路建设有限公司博深分公司"))
        # 无分支后缀的单位名原样返回（这类名称在附件内本身唯一）
        self.assertEqual(engine.gd_branch_key("深圳高速运营发展有限公司"),
                         "深圳高速运营发展有限公司")

    def test_shenzhen_g0422_attachment_holes_close_via_branch_key(self):
        """P1i：深圳 G0422 抽检段 [988,1021] 里的 3 个附件空洞靠分支名归一补齐，差额 7.131 → 0。

        附件 28 行只覆盖 25.869km（[994.599,997.340]、[1000.630,1002.534]、[1003.040,1005.526]
        三段无记录）。空洞内的明细行管养单位是「广东博大高速公路有限公司博深分公司」，
        附件里同一分支写作「广东省公路建设有限公司博深分公司」（集团）——按整串等值比 P3 落空，
        这 7.131km 两侧都不计，①表出现差额；按分支名归一后归入集团，差额归零。
        """
        # 附件真实数据（节选：3 段集团 + 尾部非集团，中间留 3 个空洞）
        index = engine.RouteCategoryIndex({("深圳", "G0422"): "高速公路"})
        index.owner_index = {("深圳", "G0422"): [
            {"start": 987.423, "end": 994.599, "owner": "集团", "keeper": "广东省公路建设有限公司博深分公司"},
            {"start": 997.340, "end": 1007.624, "owner": "集团", "keeper": "广东省公路建设有限公司博深分公司"},
            {"start": 1007.624, "end": 1021.300, "owner": "非集团", "keeper": "深圳高速运营发展有限公司"}]}
        # 明细行管养单位：同一分支但公司名不同
        keepers = [(996.0, 996.0, "广东博大高速公路有限公司博深分公司"),
                   (1001.5, 1001.5, "广东博大高速公路有限公司博深分公司"),
                   (1004.0, 1004.0, "广东博大高速公路有限公司博深分公司")]
        spans = index.owner_spans("深圳", "G0422", 988.0, 1021.0, keepers)
        covered = sum(b - a for a, b, _ in spans)
        self.assertAlmostEqual(covered, 33.0, places=6)          # 33km 全覆盖
        prov = sum(b - a for a, b, o in spans if o == "集团")
        other = sum(b - a for a, b, o in spans if o == "非集团")
        self.assertAlmostEqual(prov + other, 33.0, places=6)     # 差额 0
        self.assertAlmostEqual(prov, 19.624, places=3)           # 集团 = 12.493 + 7.131 空洞
        self.assertAlmostEqual(other, 13.376, places=3)
        # 空洞落在两侧而非整行一侧：分段里既有集团也有非集团
        self.assertEqual({o for _a, _b, o in spans}, {"集团", "非集团"})
        # 判不出的片段仍不返回（宁缺勿摊派）：管养单位在附件里跨主体时不得取多数
        ambiguous = engine.RouteCategoryIndex({("深圳", "S3"): "高速公路"})
        ambiguous.owner_index = {("深圳", "S3"): [
            {"start": 0.0, "end": 5.0, "owner": "集团", "keeper": "广东省公路建设有限公司博深分公司"},
            {"start": 5.0, "end": 10.0, "owner": "非集团", "keeper": "广东省公路建设有限公司江罗分公司"}]}
        # 片段跨两个附件分段 → 两个主体 → 判不出（不取第一段）
        self.assertIsNone(ambiguous._owner_of_slice(ambiguous.owner_index[("深圳", "S3")],
                                                    1.0, 9.0, ()))

    def test_owner_km_display_stays_additive_across_cities(self):
        """P1j：① 表三行里程用最大余数法取整，整数显示恒等（深圳 237=163+74，不是 238）。

        精确值取自 7 市产物说明句：说明句承担 3 位小数，表承担恒等。
        """
        # (全市, 非集团, 集团) → 期望的三个整数格
        for total, non_prov, prov, expected in (
            (226.818, 156.804, 70.014, ("227", "157", "70")),    # 东莞
            (173.065, 91.000, 82.065, ("173", "91", "82")),      # 中山
            (390.166, 209.941, 180.225, ("390", "210", "180")),   # 佛山
            (654.668, 541.133, 113.535, ("655", "541", "114")),   # 广州
            (417.154, 256.864, 160.290, ("417", "257", "160")),   # 惠州
            (237.283, 162.658, 74.625, ("237", "163", "74")),     # 深圳：原 163+75=238 ≠ 237
            (117.003, 97.003, 20.000, ("117", "97", "20")),       # 珠海
        ):
            with self.subTest(total=total):
                cells = engine.gd_km_split_text(total, non_prov, prov)
                self.assertEqual(cells, expected)
                self.assertEqual(int(cells[1]) + int(cells[2]), int(cells[0]))
                # 全市格是精确值普通四舍五入；分项只因补齐恒等而平移，不超过 1
                self.assertLessEqual(abs(int(cells[1]) - non_prov), 1.0)
                self.assertLessEqual(abs(int(cells[2]) - prov), 1.0)

        # 无值行不参与分配，仍写「—」；说明句的 3 位小数不受影响
        self.assertEqual(engine.gd_km_split_text(100.0, None, 60.0), ("100", "—", "60"))
        note = engine.gd_owner_gap_note(
            [{"route": "G0422", "direction": "下行", "category": "高速公路",
              "start": 0.0, "end": 33.0, "length_km": 33.0, "owner": "非集团"}], "高速公路")
        self.assertIn("33.000", note or "")       # 说明句仍是 3 位精确值

    def test_owner_km_rows_only_split_highway_owner_triple(self):
        """P1j：取整只作用于高速的「全市/非集团/集团」三行；普通国省道的全市/国道/省道原样。"""
        groups = [("全省抽检高速", None), ("深圳市抽检高速", {"km": 237.283}),
                  ("非省交通集团", {"km": 162.658}), ("省交通集团", {"km": 74.625})]
        self.assertEqual(engine.gd_km_row_texts(groups), ["", "237", "163", "74"])
        # 无主体行的普通分支：逐格原逻辑，不做恒等调整（km 为 None 的行写「—」）
        normal = [("全省抽检普通国省道", None), ("全省抽检普通国道", None), ("全省抽检普通省道", None),
                  ("珠海市抽检普通国省道", {"km": 10.4}), ("珠海市抽检普通国道", {"km": None}),
                  ("珠海市抽检普通省道", {"km": 9.6})]
        self.assertEqual(engine.gd_km_row_texts(normal), ["", "", "", "10", "—", "10"])

    def test_owner_name_alias_resolves_to_route_code(self):
        """P1h：源表「路线编号」列填路线名时按本市唯一名称反查编码（广河高速→S2），不做模糊匹配。"""
        index = engine.RouteCategoryIndex({("广州", "S2"): "高速公路"},
                                          owners={("广州", "S2"): "非集团"},
                                          names={("广州", "S2"): "广河高速"})
        self.assertEqual(index.owner_route_key("广州", "广河高速"), ("广州", "S2"))
        self.assertEqual(index.owner_for("广州", "广河高速", 2.45, 69.998), "非集团")
        # 本市不存在该名称 → 不反查，仍按原值查（判不出而非乱认）
        self.assertEqual(index.owner_route_key("广州", "沈海高速"), ("广州", "沈海高速"))
        # 同名多行时不反查（route_of_name 只收本市唯一的名称）
        ambiguous = engine.RouteCategoryIndex({("广州", "S2"): "高速公路", ("广州", "S3"): "高速公路"},
                                             names={("广州", "S2"): "广河高速", ("广州", "S3"): "广河高速"})
        self.assertNotIn(("广州", "广河高速"), ambiguous.route_of_name)

    def test_match_owner_p3_ignored_when_keeper_covers_only_part_of_span(self):
        """P1h：P3 管养单位只覆盖查询区间一段时不得收窄，否则整行被判给那一个主体（摊派）。"""
        index = engine.RouteCategoryIndex({("广州", "G0425"): "高速公路"})
        index.owner_index = {("广州", "G0425"): [
            {"start": 0.0, "end": 18.339, "owner": "非集团", "keeper": "广州公路"},
            {"start": 18.339, "end": 54.371, "owner": "集团", "keeper": "京珠北段"}]}
        # keeper 只管 0-18.339，盖不满 0-40 → 不收窄，判不出
        self.assertIsNone(engine.match_owner(index.owner_index, "广州", "G0425", 0, 40, "广州公路"))
        # keeper 覆盖整段时照常收窄（P3 原义不变）
        self.assertEqual(engine.match_owner(index.owner_index, "广州", "G0425", 0, 18, "广州公路"),
                         "非集团")
        # 附件无覆盖的尾段用 P3：管养单位在附件里唯一对应一个主体才认
        # （附件 G0425 止于 54.371，54.371 之后无覆盖）
        spans = index.owner_spans("广州", "G0425", 54.371, 60.0,
                                  [(56.0, 56.0, "京珠北段")])
        self.assertEqual(spans, [(54.371, 60.0, "集团")])
        # 尾段管养单位判不出 → 该片段不返回，差额如实保留
        self.assertEqual(index.owner_spans("广州", "G0425", 54.371, 60.0), [])

    def test_cross_owner_spans_requires_full_coverage_and_two_owners(self):
        """P1h：只拆「≥2 主体且首尾完整覆盖」的行，其余返回 None 走单值路径（零回归）。"""
        single = [(0.0, 70.0, "非集团")]
        self.assertIsNone(engine.gd_cross_owner_spans(single, 0, 70))
        # 少一段：拼不回原区间 → 不拆
        self.assertIsNone(engine.gd_cross_owner_spans([(0.0, 18.3, "非集团"), (18.3, 40.0, "集团")],
                                                      0, 41))
        self.assertIsNone(engine.gd_cross_owner_spans(None, 0, 40))
        self.assertIsNone(engine.gd_cross_owner_spans([(0.0, 18.3, "非集团"), (18.3, 40.0, "集团")],
                                                      None, 40))

    def test_measured_segments_parsed_from_detail_filenames(self):
        """P1d-b：实测段索引从明细文件名解析（K<起>K<止>，单位 km，市名取自目录名）。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "东莞-最终提交9.14"
            (root / "护栏数据明细").mkdir(parents=True)
            for name in ("S6上行K15K50-S6-交安设施现场检测-明细-20260901120000.xlsx",
                         "G4下行K2171K2151-G4-交安设施现场检测-明细-20260901120001.xlsx",
                         "S120上行中分带补测K19K25-S120-交安设施现场检测-明细-20260902105433.xlsx",
                         "S268上行中分带K63.4K64.4-S268-交安设施现场检测-明细-20260902105527.xlsx",
                         "人工自动化对比_统计结果.xlsx"):
                (root / "护栏数据明细" / name).write_bytes(b"")
            index = engine.RouteCategoryIndex({("东莞", "S6"): "高速公路", ("东莞", "G4"): "高速公路",
                                               ("东莞", "S268"): "普通国省道"})
            segments = engine.load_measured_segments(temp_dir, index)
            self.assertEqual(segments[("东莞", "S6")], [(15.0, 50.0)])
            self.assertEqual(segments[("东莞", "G4")], [(2151.0, 2171.0)])
            self.assertEqual(segments[("东莞", "S120")], [(19.0, 25.0)])
            self.assertEqual(segments[("东莞", "S268")], [(63.4, 64.4)])
            self.assertNotIn(("东莞", "人工自动化对比_统计结果"), segments)
            # 目录不存在/无市名 → 空 dict，不抛异常
            self.assertEqual(engine.load_measured_segments(Path(temp_dir) / "不存在", index), {})


class GuangdongImportSourceTests(unittest.TestCase):
    """导入源识别口径：标线取“标线统计”表（区间统计/单路线明细为派生物，不得重复读取），
    护栏高度与螺栓取“护栏数据明细”文件夹；按市存放时每市各命中一份。"""

    @staticmethod
    def _tree(temp_dir: Path) -> Path:
        root = Path(temp_dir) / "最终复核数据"
        for city, marking_dir, guardrail_dir in (
            ("东莞-最终提交9.14", "东莞标线数据明细", "东莞护栏数据明细"),
            ("广州-最终提交9.15", "广州标线数据明细", "广东护栏数据明细"),   # 市名与文件夹名不一致
            ("中山-最终提交9.14", "中山统计标线数据", "中山市护栏数据明细"),  # 标线目录命名不同
        ):
            m = root / city / marking_dir
            m.mkdir(parents=True)
            (m / "标线统计.xlsx").write_bytes(b"")
            (m / "标线区间统计.xlsx").write_bytes(b"")      # 派生物：不得被当作数据源
            (m / "单路线明细").mkdir()
            (m / "单路线明细" / "S47下行K155K130标线2.csv").write_text("x")
            g = root / city / guardrail_dir
            g.mkdir(parents=True)
            (g / "G0422下行K1005K965-交安设施现场检测-明细.xlsx").write_bytes(b"")
            for stage in ("进场前人工复核数据对比", "进场后人工复核数据对比"):
                (root / city / stage).mkdir()
                (root / city / stage / f"人工自动化对比{city[:2]}.xlsx").write_bytes(b"")
        return root

    def test_marking_uses_statistics_table_and_guardrail_uses_detail_folder(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            root = self._tree(temp_dir)
            marking, guardrail = engine.detect_guangdong_sources(root)

        self.assertEqual(len(marking), 3, f"每市各一份标线统计表：{[p.name for p in marking]}")
        self.assertTrue(all(p.name == "标线统计.xlsx" for p in marking))
        # 区间统计与单路线明细不得入选
        self.assertFalse([p for p in marking if "区间" in p.name or "单路线明细" in p.parts])
        self.assertEqual(len(guardrail), 3, f"每市各一份护栏明细文件夹：{[p.name for p in guardrail]}")
        self.assertTrue(all("护栏数据明细" in p.name for p in guardrail))
        # 市名与文件夹名不一致（广州/广东）、标线目录命名不同（中山统计标线数据）都要命中
        self.assertEqual({p.parent.parent.name for p in marking},
                         {"东莞-最终提交9.14", "广州-最终提交9.15", "中山-最终提交9.14"})

    def test_single_city_folder_and_file_root_scan(self):
        """单市文件夹只命中一份；扫描器需支持直接指向单个文件（标线统计表）。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = self._tree(temp_dir)
            marking, guardrail = engine.detect_guangdong_sources(root / "东莞-最终提交9.14")

            self.assertEqual(len(marking), 1)
            self.assertEqual(len(guardrail), 1)
            self.assertEqual(guardrail[0].name, "东莞护栏数据明细")
            # 文件根：scan() 必须接受单个文件而不是抛“项目资料文件夹不存在”
            scanner = engine.GuangdongInputScanner(marking[0], None, log=lambda _m: None)
            self.assertEqual(scanner.scan()["marking"], [])

    def test_marking_keeps_only_the_files_own_city(self):
        """跨市路段只保留本市那份，且“韶关市/韶关”两种写法归一后合并计数。"""
        rows = ([{"city": "东莞", "route": "G220", "segment": "1-2"}] * 5
                + [{"city": "深圳", "route": "G0422", "segment": "1005+000-1004+980"}] * 3
                + [{"city": "惠州", "route": "G220", "segment": "3-4"}] * 1)

        own, kept = engine.filter_own_city_marking(rows)

        self.assertEqual(own, "东莞")
        self.assertEqual(len(kept), 5)
        self.assertTrue(all(r["city"] == "东莞" for r in kept))
        # 写法差异不得被当成两个市
        own2, kept2 = engine.filter_own_city_marking(
            [{"city": "韶关市"}] * 4 + [{"city": "韶关"}] * 3 + [{"city": "清远市"}] * 2)
        self.assertEqual(own2, "韶关")
        self.assertEqual(len(kept2), 7)
        # 空输入不得抛异常
        self.assertEqual(engine.filter_own_city_marking([]), (None, []))

    def test_guardrail_dir_name_variants_including_nested(self):
        """护栏目录实测命名有 5 种形态（含嵌套在「云浮明细」这类父目录下），都不得漏扫。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "最终复核数据"
            for city, guardrail_rel in (
                ("云浮9.9", "云浮明细/2. 护栏高度、螺栓数据"),   # 父目录名不含「护栏」，只能靠子目录
                ("清远9.10", "清远市明细/2. 护栏高度、螺栓数据"),
                ("湛江9.10", "2. 护栏螺栓、高度数据"),            # 螺栓在前
                ("茂名9.10", "护栏明细数据"),                      # 字序与常量相反
                ("阳江9.11", "护栏明细"),                          # 只有「护栏明细」
            ):
                (root / city / guardrail_rel).mkdir(parents=True)

            for city_dir in sorted(p for p in root.iterdir() if p.is_dir()):
                marking, guardrail = engine.detect_guangdong_sources(city_dir)

                self.assertEqual(len(guardrail), 1, f"{city_dir.name} 护栏目录未命中：{guardrail}")

    def test_marking_hit_does_not_short_circuit_guardrail_fallback(self):
        """标线命名命中不得短路护栏回退：有标线、护栏目录名未命中时护栏仍须按表头找到（否则高度/螺栓静默为 0）。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "云浮9.9"
            marking_dir = root / "云浮明细" / "1. 云浮标线处理数据"
            marking_dir.mkdir(parents=True)
            (marking_dir / "标线统计.xlsx").write_bytes(b"")
            guardrail_dir = root / "云浮明细" / "2. 交安数据"   # 目录名不含「护栏」，命名规则必然不命中
            guardrail_dir.mkdir(parents=True)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["桩号", "护栏类型", "梁板中心高度(mm)", "地市", "路线编号", "检测方向"])
            ws.append(["1000+000-1000+020", "波形梁", 750, "云浮市", "S368", "上行"])
            wb.save(guardrail_dir / "S368上行K1000K999-交安设施现场检测-明细.xlsx")

            marking, guardrail = engine.detect_guangdong_sources(root)

        self.assertEqual([p.name for p in marking], ["标线统计.xlsx"])
        self.assertEqual([p.name for p in guardrail], ["2. 交安数据"])

    def test_guardrail_empty_but_header_files_present_must_warn(self):
        """禁止静默 0：高度/螺栓解析为 0 而项目内确有护栏表头明细时必须给出含目录的警告。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "云浮9.9"
            guardrail_dir = root / "2. 护栏高度、螺栓数据"
            guardrail_dir.mkdir(parents=True)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["桩号", "护栏类型", "梁板中心高度(mm)", "地市", "路线编号", "检测方向"])
            ws.append(["1000+000-1000+020", "波形梁", 750, "云浮市", "S368", "上行"])
            wb.save(guardrail_dir / "S368上行K1000K999-交安设施现场检测-明细.xlsx")

            empty = {"height": [], "bolt": [], "marking": [{"city": "云浮市"}]}
            warning = engine.guardrail_missing_warning(root, empty)

        self.assertIsNotNone(warning, "护栏解析为 0 且存在护栏表头文件时不得静默")
        self.assertIn("2. 护栏高度、螺栓数据", warning or "")
        # 有数据时不得产生噪音警告
        self.assertIsNone(engine.guardrail_missing_warning(
            Path(temp_dir), {"height": [{"city": "云浮市"}], "bolt": [], "marking": []}))

    def test_naming_miss_falls_back_to_header_detection(self):
        """命名未命中（新格式）时回退表头识别，不得直接报空。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir) / "旧格式项目"
            d = root / "含糊命名目录"
            d.mkdir(parents=True)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["桩号", "逆反射亮度系数", "目标值", "地市", "路线编号", "检测方向", "管养单位"])
            ws.append(["1000+000-1000+020", 90, 80, "云浮市", "S368", "上行", "云浮中心"])
            wb.save(d / "随便什么名字.xlsx")

            marking, guardrail = engine.detect_guangdong_sources(root)

        # 命名未命中时回退为目录级表头识别（旧版行为），要求非空且指向含数据的目录
        self.assertEqual([p.name for p in marking], ["含糊命名目录"])
        self.assertEqual(guardrail, [])


class R2FolderCityTests(unittest.TestCase):
    """R2（用户明确要求）：城市归属以数据所在文件夹定市。

    文件夹内记录即使“地市”列写的是别的市，也计入该文件夹对应的市；地市列不再决定归属、
    也不再因此丢弃。旧口径按“文件内多数地市”过滤（东莞实测丢弃 2000 条标线记录）。
    """

    HEADERS = ["地市", "路线编号", "检测方向", "管养单位", "桩号",
               "主车道左侧标线逆反亮度系数", "主车道右侧标线逆反亮度系数",
               "左侧标线逆反射目标值", "右侧标线逆反射目标值", "计算区间"]

    @classmethod
    def _marking_row(cls, city, route, station):
        return dict(zip(cls.HEADERS, [city, route, "下行", "甲管养单位", station, 40, 90, 50, 80, 20]))

    @classmethod
    def _write_marking(cls, path, rows):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(cls.HEADERS)
        for row in rows:
            ws.append([row.get(key) for key in cls.HEADERS])
        wb.save(path)
        wb.close()

    def test_folder_city_is_authoritative_for_records_inside_it(self):
        """文件夹内的记录一律记为文件夹市；旧口径丢弃的“其他市标记”记录必须保留。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir) / "东莞-最终提交9.14" / "东莞标线数据明细"
            folder.mkdir(parents=True)
            path = folder / "标线统计.xlsx"
            rows = ([self._marking_row("东莞", "G220", f"1000+{i:03d}.0-1000+{i + 20:03d}.0") for i in (0, 20, 40)]
                    + [self._marking_row("深圳", "G0422", f"1005+{i:03d}.0-1005+{i + 20:03d}.0") for i in (0, 20)])
            self._write_marking(path, rows)
            route_index = engine.RouteCategoryIndex({("东莞", "G220"): "普通国省道", ("深圳", "G0422"): "高速公路"})
            logs = []

            scanner = engine.GuangdongInputScanner(path, route_index, log=logs.append)
            self.assertEqual(scanner.default_city, "东莞")   # 文件夹定市：数据在 东莞-最终提交9.14 下
            scanned = scanner.scan()
            # 5 行 × 左右两侧 = 10 条，其中 4 条“地市”列写深圳
            self.assertEqual(len(scanned["marking"]), 10)
            self.assertEqual({row["city"] for row in scanned["marking"]}, {"东莞"})
            kept = engine.own_city_marking(scanned["marking"], logs.append, scanner.default_city)
            bundles = engine.GuangdongBatchRunner.build_bundles({"marking": kept, "issues": []}, route_index)

        self.assertEqual(len(kept), 10)                      # 旧口径只保留 6 条
        self.assertEqual(sum(1 for row in kept if row.get("declared_city")), 4)
        self.assertIn("按文件夹归属东莞取数，保留其他市标记记录 4 条",
                      [line.strip() for line in logs])
        # 显示层市名带「市」（P1d-b 裁决 1）；旧口径会把 4 条归到“深圳”分组
        self.assertEqual(list(bundles), ["东莞市"])
        self.assertEqual(len(bundles["东莞市"]["marking"]), 10)
        # G0422 归属深圳，P1d-b 起东莞不再借用其类别 → 只有东莞自己的 G220 有类别
        self.assertEqual({row["category"] for row in bundles["东莞市"]["marking"]},
                         {"普通国省道", None})

    def test_missing_route_category_falls_back_and_never_aborts_city(self):
        """路线不在本市路线表：P1d-b 起不再跨市借类别，留空并记 issues（整市报告不中断）。"""
        route_index = engine.RouteCategoryIndex({
            ("东莞", "G220"): "普通国省道",
            ("深圳", "G0422"): "高速公路",      # 归属深圳：P1d-b 起东莞不再借用（P0-A）
            ("广州", "S999"): "高速公路",
            ("佛山", "S999"): "普通国省道",     # 同号跨市且类别冲突 → 无法唯一判定
        })
        scanned = {"issues": [], "marking": [
            {"city": "东莞", "route": "G0422", "direction": "下行", "segment": "a", "station_m": 1000.0},
            {"city": "东莞", "route": "S999", "direction": "下行", "segment": "b", "station_m": 2000.0},
        ]}

        bundle = engine.GuangdongBatchRunner.build_bundles(scanned, route_index)["东莞市"]

        self.assertNotIn("_error", bundle)                   # 旧实现抛 KeyError → 整市报告中断
        by_route = {row["route"]: row for row in bundle["marking"]}
        self.assertIsNone(by_route["G0422"]["category"])      # 他市路线码：不借类别（P0-A）
        self.assertIsNone(by_route["S999"]["category"])
        self.assertEqual(by_route["S999"]["city"], "东莞市")   # 仍归属该文件夹的市（显示层带「市」）
        self.assertTrue(any("S999" in message for message in bundle["issues"]),
                        f"无法判定类别的行必须记入 issues：{bundle['issues']}")
        self.assertEqual(bundle["unknown_categories"],
                         [{"kind": "marking", "route": "G0422", "count": 1},
                          {"kind": "marking", "route": "S999", "count": 1}])


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
        # “共X页”只统计目录以后的页数：正文节用 SECTIONPAGES，不再用全文 NUMPAGES。
        self.assertIn("SECTIONPAGES", footers[body_footer])
        self.assertNotIn("NUMPAGES", footers[body_footer])
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

    def test_bolt_example_title_shows_electronic_station_then_missing_then_raw(self) -> None:
        """螺栓示例标题＝电子（修正）桩号＋螺栓缺失X颗＋原始桩号，两个桩号均保留一位小数。"""
        from openpyxl.drawing.image import Image as XLImage

        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            summary = root / "summary.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "各区县项目概况"
            sheet.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
            sheet.append([1, "万州区", "G210", "", "一级", 2264.0, 2265.0, 1.0, 1.0])
            workbook.save(summary)

            detail = root / "detail"
            detail.mkdir()
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "螺栓明细"
            sheet.append(["路线编号", "方向", "原始桩号", "电子修正桩号", "拼接螺栓数量（颗）",
                          "拼接螺栓缺失数量（颗）", "连接螺栓数量（颗）", "连接螺栓缺失数量（颗）"])
            sheet.append(["G210", "上行", "K2264+409.0", "K2264+386", 10, 2, 10, 0])
            workbook.save(detail / "重庆市-万州区-交安设施现场检测-明细.xlsx")

            png = base64.b64decode(
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
            )
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.append(["序号", "路线", "方向", "原始桩号", "电子修正桩号", "病害类型", "工程量", "单位", "病害照片"])
            sheet.append([1, "G210", "上行", "K2264+409.0", "K2264+386", "波形护栏螺栓缺失", 2, "颗", None])
            sheet.add_image(XLImage(io.BytesIO(png)), "I2")
            workbook.save(detail / "重庆市-万州区-交安设施现场检测-病害清单.xlsx")

            template = root / "template.md"
            template.write_text("# 报告\n\n<!-- inject:bolt -->\n", encoding="utf-8")
            output = root / "output"
            config = engine.Config(root, summary, detail, template, output, disease_dir=detail)
            engine.generate_statistics_and_report(
                config, log=lambda _: None, process_height=False,
                process_bolts=True, process_tci=False, require_template=False,
            )
            docx = next(output.rglob("*.docx"))
            with zipfile.ZipFile(docx) as archive:
                document = archive.read("word/document.xml").decode("utf-8")
        self.assertIn("K2264+386.0 螺栓缺失2颗 K2264+409.0", document)


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

    def test_height_excel_writes_tolerance_comparison_sheets_without_charts(self) -> None:
        """高度统计工作簿不再画图，改为写出 ±50mm / ±40mm 两组分档对照表。"""
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            summary = root / "summary.xlsx"
            self._write_summary(summary)
            segments = engine.read_segments(summary)
            county_segments = [s for s in segments if s["county"] == "万州区"]

            def record(station, height):
                return {
                    "file": "f.xlsx", "direction": "上行", "station": float(station),
                    "raw_station": float(station), "electronic_station": float(station),
                    "basis": "电子修正桩号", "kind": "二波", "height": float(height), "segment": 0,
                }

            # 600mm 落在三段内；555mm 只在 ±50mm 内（560mm 以下），用于区分两组对照。
            records = [record(2264100, 600), record(2264200, 555)]
            height_stats = engine.make_stats(county_segments, records)
            config = engine.Config(
                root, summary, root, root / "template.md", root / "output", county="万州区",
            )
            engine.make_excel(config, county_segments, height_stats=height_stats, height_records=records)
            workbook = openpyxl.load_workbook(config.out_xlsx)
            sheets = workbook.sheetnames
            charts = [len(sheet._charts) for sheet in workbook.worksheets]
            main = workbook["二波统计"]
            main_pass = main.cell(2, [cell.value for cell in main[1]].index("合格率") + 1).value
            labels50 = [cell.value for cell in workbook["二波对照（±50mm）"][1]]
            labels40 = [cell.value for cell in workbook["二波对照（±40mm）"][1]]
            comp50 = workbook["二波对照（±50mm）"]
            comp40 = workbook["二波对照（±40mm）"]
            counts = (
                comp50.cell(2, labels50.index("检测点数") + 1).value,
                comp40.cell(2, labels40.index("检测点数") + 1).value,
            )
            pass50 = comp50.cell(2, labels50.index("合格率（±50mm）") + 1).value
            pass40 = comp40.cell(2, labels40.index("合格率（±40mm）") + 1).value
            workbook.close()
        self.assertIn("二波对照（±50mm）", sheets)
        self.assertIn("二波对照（±40mm）", sheets)
        self.assertEqual(charts, [0] * len(sheets))
        self.assertEqual(main_pass, 0.5)
        self.assertEqual(counts, (2, 2))
        self.assertEqual(labels50[9:14], ["h＜550", "550≤h＜580", "580≤h≤620", "620＜h≤650", "h＞650"])
        self.assertEqual(labels40[9:14], ["h＜560", "560≤h＜580", "580≤h≤620", "620＜h≤640", "h＞640"])
        self.assertEqual(pass50, 1.0)
        self.assertEqual(pass40, 0.5)


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
                [minimal_docx._strip_section_number(h) for h in h1 if not h.startswith("附表")],
                ["概况", "组织实施情况", "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测结果", "波形梁护栏螺栓缺失", "结论与建议"],
            )
        for required in ("检测依据", "检测设备与评定方法", "沿线设施技术状况评价", "波形梁护栏横梁中心高度检测", "波形梁护栏螺栓缺失检测"):
            self.assertTrue(any(required in h for h in headings), f"missing static heading: {required}")
        # 按路线合并后，部分缺源不再建立独立区段小节。
        self.assertIn("部分检测区间未提供TCI数据", text)
        self.assertIn("共检出拼接螺栓80颗，连接螺栓120颗，缺失螺栓4颗，缺失率为1.96%", text)

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

    def test_height_example_points_are_filtered_to_photo_rows_before_selection(self) -> None:
        """选点前先排除无照片点；一个可挂图点都没有时返回空（不输出无图示例表）。"""
        records = [
            {"height": 560.0 + index, "station": 1000.0 + index, "raw_station": 1000.0 + index,
             "electronic_station": 1000.0 + index, "direction": "上行", "route": "G210", "county": "甲县"}
            for index in range(10)
        ]
        photo = {"role": "height", "route": "G210", "county": "", "workbook": Path("disease.xlsx"), "media": "xl/media/image1.png"}
        index = {("上行", 1003.0): [dict(photo)], ("上行", 1006.0): [dict(photo)]}
        points = engine.select_height_example_points(records, 0, "二波", photo_index=index)
        self.assertEqual([point["station"] for point in points], [1006.0])
        for point in points:
            self.assertTrue(engine.matching_height_photos(point, index), point)
        self.assertEqual(engine.select_height_example_points(records, 0, "二波", photo_index={("上行", 9999.0): [dict(photo)]}), [])
        self.assertEqual(len(engine.select_height_example_points(records, 0, "二波")), 1)

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
        self.assertEqual(set(segment_photos), {(0, "上行")})

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


class ChongqingTemplateUpdateTests(unittest.TestCase):
    """R13：4.x 表列、附表按路线拆分、页码“共X页”、Excel 区县合计行。"""

    @staticmethod
    def _segments():
        return [dict(county="甲县", route="G210", start=1000.0, end=2000.0, mileage=1.0, grade="二级公路", manager=""),
                dict(county="甲县", route="G319", start=3000.0, end=4000.0, mileage=1.0, grade="二级公路", manager="")]

    @staticmethod
    def _height_records(segments):
        records = []
        for index, segment in enumerate(segments):
            for kind, heights in (("二波", (600.0, 555.0)), ("三波", (690.0, 650.0))):
                for offset, height in enumerate(heights, 1):
                    station = segment["start"] + offset * 10
                    records.append(dict(file="t.xlsx", direction="上行", station=station, raw_station=station,
                                        electronic_station=station, basis="电子修正桩号", kind=kind,
                                        height=height, segment=index))
        return records

    def test_height_route_table_drops_valid_and_pass_point_columns(self) -> None:
        segments = self._segments()
        records = self._height_records(segments)
        with tempfile.TemporaryDirectory() as temp_dir:
            doc = Document()
            minimal_docx._section_height(doc, segments, engine.make_stats(segments, records), records, {}, Path(temp_dir))
        route_tables = [table for table in doc.tables if "方向" in [cell.text for cell in table.rows[0].cells]]
        self.assertTrue(route_tables)
        for table in route_tables:
            headers = [cell.text for cell in table.rows[0].cells]
            self.assertNotIn("有效点数", headers)
            self.assertNotIn("合格点数", headers)
            self.assertIn("合格率（%）", headers)
            self.assertEqual(len(headers), len(table.rows[1].cells))
            self.assertTrue(all(len(row.cells) == len(headers) for row in table.rows))
        self.assertEqual([cell.text for cell in route_tables[0].rows[0].cells][:5],
                         ["序号", "方向", "起点桩号", "止点桩号", "合格率（%）"])

    def test_appendix_keeps_one_table_with_route_blocks(self) -> None:
        headers = ["序号", "路线编号", "区县", "方向", "起点桩号", "止点桩号", "TCI", "等级"]
        segments = [dict(county="甲县", route=route, start=start, end=end, mileage=(end - start) / 1000)
                    for route, start, end in [("G210", 1000.0, 2000.0), ("G210", 3000.0, 4000.0),
                                              ("G319", 5000.0, 6000.0)]]
        records = [dict(segment=index, direction="上行", station=segment["start"] + 10,
                        light=0, heavy=0, sign=1, marking=0)
                   for index, segment in enumerate(segments)]
        stats = engine.make_tci_stats(segments, records)
        doc = Document()
        minimal_docx._section_tci_appendix(doc, segments, stats)
        titles = [p.text for p in doc.paragraphs if p.text.startswith("附表")]
        self.assertEqual(titles, ["附表1 重庆市甲县交安设施技术状况评定明细"])
        self.assertEqual(doc._toc_entries, [("", titles[0])])
        self.assertTrue(all(p.style.name == "Heading 1" for p in doc.paragraphs if p.text.startswith("附表")))
        self.assertEqual(len(doc.tables), 1)
        table = doc.tables[0]
        # 序号按路线重排，块间第 4 行为合并空白行
        self.assertEqual([row.cells[0].text for row in table.rows],
                         ["序号", "1", "2", "3", "", "序号", "1", "2"])
        for index in (0, 5):  # 表首表头 + 每条路线前的重复表头
            cells = table.rows[index].cells
            self.assertEqual([cell.text for cell in cells], headers)
            self.assertTrue(cells[0].paragraphs[0].runs[0].bold)
            self.assertIn("w:shd", table.rows[index]._tr.xml)
        separator = table.rows[4]._tr.findall(qn("w:tc"))
        self.assertEqual(len(separator), 1)
        self.assertEqual(separator[0].find(qn("w:tcPr")).find(qn("w:gridSpan")).get(qn("w:val")), "8")
        self.assertEqual(table.rows[1].cells[1].text, "G210")
        self.assertEqual(table.rows[6].cells[1].text, "G319")
        self.assertTrue(table.rows[0]._tr.xpath("./w:trPr/w:tblHeader"))

    def test_body_footer_total_pages_and_jiuhao_font(self) -> None:
        root = Path(__file__).parent / "artifacts" / "chongqing-footer"
        root.mkdir(parents=True, exist_ok=True)
        template = root / "template.md"
        template.write_text(
            "# @cover 主标题行一|主标题行二|重庆市|报告编号：BG-2026-T9"
            "|项目名称：测试项目|委托单位：测试单位|测试公司|二〇二六年七月\n"
            "\n<!-- toc -->\n\n# 1 概况\n",
            encoding="utf-8",
        )
        output = root / "output"
        output.mkdir(exist_ok=True)
        config = engine.Config(root, root / "summary.xlsx", root, template, output)
        minimal_docx.make_report(config, [], None, [], None, [], None, [], {}, root, skeleton_md=template)
        doc = Document(Path(output) / engine.OUT_DOCX_NAME)
        body_footer = doc.sections[-1].footer
        paragraph = body_footer.paragraphs[0]
        xml = paragraph._p.xml
        self.assertIn("SECTIONPAGES", xml)
        self.assertNotIn("NUMPAGES", xml)
        self.assertEqual("".join(run.text for run in paragraph.runs), "第  页 共  页")
        self.assertTrue(paragraph.runs)
        for run in paragraph.runs:  # 小五号＝9pt，含页码域内的数字
            self.assertIsNotNone(run.font.size, "页脚 run 必须显式设字号")
            self.assertEqual(run.font.size.pt, 9.0)

    def test_excel_appends_county_total_rows(self) -> None:
        with tempfile.TemporaryDirectory() as temp_dir:
            root = Path(temp_dir)
            summary = root / "summary.xlsx"
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "各区县项目概况"
            sheet.append(["序号", "区县", "路线编号", "路线名", "公路等级", "起点桩号", "止点桩号", "里程", "总里程"])
            sheet.append([1, "甲县", "G210", "", "二级", 1.0, 2.0, 1.0, 2.0])
            sheet.append([2, "甲县", "G319", "", "二级", 3.0, 4.0, 1.0, 2.0])
            workbook.save(summary)
            segments = engine.read_segments(summary)
            heights = self._height_records(segments)
            bolts = [dict(segment=index, direction="上行", station=segment["start"] + 10,
                          raw_station=segment["start"] + 10, file="t.xlsx", basis="电子修正桩号",
                          splice=10, connection=5, splice_missing=1, connection_missing=0)
                     for index, segment in enumerate(segments)]
            tci_records = [dict(segment=index, direction="上行", station=segment["start"] + 10,
                                light=0, heavy=0, sign=1, marking=0)
                           for index, segment in enumerate(segments)]
            tci_stats = engine.make_tci_stats(segments, tci_records)
            config = engine.Config(root, summary, root, root / "template.md", root / "output", county="甲县")
            engine.make_excel(config, segments, height_stats=engine.make_stats(segments, heights),
                              height_records=heights, bolt_stats=engine.make_bolt_stats(segments, bolts),
                              bolt_records=bolts, tci_stats=tci_stats, tci_records=tci_records)
            sheets = openpyxl.load_workbook(config.out_xlsx)
            try:
                for tolerance in (50, 40):
                    for kind in ("二波", "三波"):
                        sheet = sheets[f"{kind}对照（±{tolerance}mm）"]
                        expected = engine._height_type_bins(heights, kind, tolerance, pass_bins=(1, 2, 3))
                        self.assertEqual(sheet.max_row, 4)  # 表头+2条路线+合计
                        self.assertEqual(sheet.cell(4, 1).value, "合计")
                        self.assertEqual(sheet.cell(4, 2).value, "甲县")
                        self.assertEqual(sheet.cell(4, 3).value, "全部路线")
                        headers = [cell.value for cell in sheet[1]]
                        self.assertEqual(sheet.cell(4, headers.index("检测点数") + 1).value, expected["count"])
                        self.assertEqual(sheet.cell(4, headers.index(f"合格率（±{tolerance}mm）") + 1).value,
                                         expected["pass"] / 100)
                        self.assertEqual([round(sheet.cell(4, headers.index(label) + 1).value * 100, 2)
                                          for label in engine.height_bin_labels(kind, tolerance)], expected["pcts"])
                # 主统计表不加合计行（用户只要求对照表/螺栓/TCI）
                self.assertEqual(sheets["二波统计"].max_row, 3)
                bolt_sheet = sheets["螺栓缺失统计"]
                total = engine.merge_bolt_totals(engine.make_bolt_stats(segments, bolts))
                self.assertEqual(bolt_sheet.max_row, 4)
                self.assertEqual(bolt_sheet.cell(4, 1).value, "合计")
                self.assertEqual(bolt_sheet.cell(4, 3).value, "全部路线")
                self.assertEqual(bolt_sheet.cell(4, 7).value, total["splice"])
                self.assertEqual(bolt_sheet.cell(4, 8).value, total["connection"])
                self.assertEqual(bolt_sheet.cell(4, 9).value, total["missing"])
                self.assertAlmostEqual(bolt_sheet.cell(4, 10).value, total["rate"] / 100, places=10)
                units = [unit for stat in tci_stats for unit in stat["units"]]
                expected_tci = sum(unit["tci"] for unit in units) / len(units)
                tci_sheet = sheets["沿线设施统计"]
                self.assertEqual(tci_sheet.max_row, 4)
                self.assertEqual(tci_sheet.cell(4, 1).value, "合计")
                self.assertEqual(tci_sheet.cell(4, 3).value, "全部路线")
                self.assertEqual(tci_sheet.cell(4, 5).value, engine.format_station(segments[-1]["end"]))
                self.assertAlmostEqual(tci_sheet.cell(4, 11).value, expected_tci, places=6)
                self.assertEqual(tci_sheet.cell(4, 12).value, engine.tci_grade(expected_tci))
                notes = [row[0].value for row in sheets["统计说明"].iter_rows()]
                self.assertIn("合计行", notes)
            finally:
                sheets.close()


if __name__ == "__main__":
    unittest.main()
class T2fFixesTests(unittest.TestCase):
    """T2f：标线颜色映射 / 引言逐字与上标 / 分页标记 / 管养单位显示名 / 合格值公差。"""

    @staticmethod
    def _bundle() -> dict:
        marking = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K2+000",
                    "marking_position": "左侧标线", "value": 80.0, "target": 80.0, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        detail = [
            {"indicator": "marking", "category": "高速公路", "route": "S14", "direction": "上行",
             "segment": "K1+000～K2+000", "city": "测试市", "manual": 78.0, "automatic": 80.0,
             "marking_position": "标线1"},
            {"indicator": "height", "category": "高速公路", "route": "S14", "direction": "上行",
             "segment": "K1+500～K2+000", "city": "测试市", "manual": 600.0, "automatic": 604.0,
             "gtype": "两波护栏"},
        ]
        return {"city": "测试市", "marking": marking, "height": [], "bolt": [], "notes": [],
                "comparison_detail": detail, "weak_segments": [], "route_segments": []}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_marking_color_derived_from_target(self) -> None:
        """P1-1：颜色由自动化单元目标值推导（80→白色、50→黄色），判定不了写 —，不回填道路等级。"""
        auto_white = [{"route": "G1", "direction": "上行", "value": 80.0, "target": 80.0}]
        auto_yellow = [{"route": "G1", "direction": "上行", "value": 50.0, "target": 50.0}]
        row = {"route": "G1", "direction": "上行", "automatic": 80.0, "category": "高速公路"}
        self.assertEqual(engine.gd_marking_color(row, auto_white), "白色")
        self.assertEqual(engine.gd_marking_color({**row, "automatic": 50.0}, auto_yellow), "黄色")
        # 测值命中单元优先：命中即只认该单元的目标值
        self.assertEqual(engine.gd_marking_color(row, auto_white + auto_yellow), "白色")
        # 测值命不中、退化为该路线方向全部单元且目标值不唯一 → 无法判定
        miss = {**row, "automatic": 999.0}
        self.assertEqual(engine.gd_marking_color(miss, auto_white + auto_yellow), "—")
        # 无自动化单元 → —
        self.assertEqual(engine.gd_marking_color(row, []), "—")

    def test_intro_paragraph_verbatim_and_superscript(self) -> None:
        """P1-2：引言段按模板-1 逐字（含 ±20mm 与 JTG 5210—2018），-²/-¹ 为上标 run（模板 4 处）。"""
        document = self._write()
        intro = next(p for p in document.paragraphs if "mcd·m" in p.text)
        text = intro.text.strip()
        self.assertIn("JTG 5210—2018", text)
        self.assertIn("两波形梁钢护栏为600mm±20mm、三波形梁钢护栏为697mm±20mm即为合格", text)
        self.assertIn("mcd·m-²·lx-¹", text)
        self.assertNotIn("合格合格率", text)
        self.assertNotIn("重新统计", text)
        self.assertEqual(sum(1 for run in intro.runs if run.font.superscript), 4)

    def test_pagination_flags(self) -> None:
        """P2-2/P2-3/P3-6：表格行 cantSplit、题注 keep_with_next、正文引导段 keep_with_next。"""
        document = self._write()
        self.assertGreater(document.element.xml.count("cantSplit"), 0)
        captions = [p for p in document.paragraphs if p.style.name == "Caption"]
        self.assertTrue(captions)
        self.assertTrue(all(p.paragraph_format.keep_with_next for p in captions))
        lead = next(p for p in document.paragraphs if p.text.strip() == "（1）分类开展显性问题突击整治")
        self.assertTrue(lead.paragraph_format.keep_with_next)

    def test_manager_display_keeps_route_table_name(self) -> None:
        """P2-5：显示名统一到路线分类表口径（不再简写公司/去地市前缀），缺值写 —。"""
        self.assertEqual(engine.manager_display("东莞市公路事务中心", "东莞"), "东莞市公路事务中心")
        self.assertEqual(engine.manager_display("广东博大高速公路有限公司博深分公司"), "广东博大高速公路有限公司博深分公司")
        self.assertEqual(engine.manager_display(None, "东莞"), "—")

    def test_height_grade_with_tolerance(self) -> None:
        """P3-1：高度人工复核表 `合格值(mm)` 按模板带公差。"""
        document = self._write()
        table = next(t for t in document.tables if any("合格值" in c.text for c in t.rows[0].cells))
        column = next(i for i, c in enumerate(table.rows[0].cells) if "合格值" in c.text)
        values = {row.cells[column].text.strip() for row in table.rows[1:]} - {"合格值(mm)"}
        self.assertEqual(values, {"600±20mm"})
class T2gRoundThreeTests(unittest.TestCase):
    """T2g（用户第 3 轮）：① 只留 1 段 / 三小节带序号 / 典型段标题为正文级且补「段」。"""

    @staticmethod
    def _bundle() -> dict:
        marking = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                    "marking_position": "左侧标线", "value": 60.0, "target": 80.0, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        height = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                   "guardrail_type": "两波", "height": 600.0, "city": "测试市",
                   "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        return {"city": "测试市", "marking": marking, "height": height, "bolt": [], "notes": [],
                "comparison_detail": [], "weak_segments": [], "route_segments": []}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_overall_section_has_single_paragraph(self) -> None:
        """G1：① 总体情况 下只剩 1 段（统计口径/检测点/应安装螺栓段已删）。"""
        document = self._write()
        texts = [p.text.strip() for p in document.paragraphs]
        joined = "\n".join(texts)
        # 三处被删的口径段特征串（唯一出现在 ① 第 2 段）
        self.assertNotIn("有效计算单元", joined)
        self.assertNotIn("覆盖整公里检测段", joined)
        self.assertNotIn("本次共检测应安装螺栓", joined)
        for index, text in enumerate(texts):
            if text != "①总体情况":
                continue
            # ① 下恰好 1 段正文：其后第一段不得是统计口径段
            body = texts[index + 1]
            for token in ("本次共抽检", "本次共获得有效检测点", "本次共检测应安装螺栓"):
                self.assertFalse(body.startswith(token), body)

    def test_auto_sections_are_numbered(self) -> None:
        """G2：三个自动化小节标题带 `（1）（2）（3）`，Heading 4 + outlineLvl=3 + 仿宋_GB2312 12 bold。"""
        document = self._write()
        by_text = {p.text.strip(): p for p in document.paragraphs}
        for title, level in (("（1）标线逆反射亮度系数情况", 4), ("（2）波形梁护栏中心高度情况", 4),
                             ("（3）波形梁护栏螺栓缺失情况", 4)):
            paragraph = by_text[title]
            self.assertEqual(paragraph.style.name, "Heading %d" % level)
            self.assertEqual(engine.GuangdongChapterWriter._outline_level(paragraph), level)
            run = paragraph.runs[0]
            self.assertEqual(run.font.name, "Times New Roman")
            self.assertEqual(run.font.size.pt, 12.0)
            self.assertTrue(run.bold)
            self.assertEqual(run._element.rPr.rFonts.get(qn("w:eastAsia")), "仿宋_GB2312")

    def test_typical_segment_heading_is_body_level_with_suffix(self) -> None:
        """G3：`1）…` 典型段标题为正文级（Normal、无 outlineLvl）、无「（约N公里）」、末尾带「段」、楷体 12 bold。"""
        document = self._write()
        items = [p for p in document.paragraphs if re.match(r"^\d+）", p.text.strip())]
        self.assertTrue(items, "样例数据应生成 ③ 典型段标题")
        for paragraph in items:
            text = paragraph.text.strip()
            self.assertEqual(paragraph.style.name, "Normal")
            self.assertIsNone(engine.GuangdongChapterWriter._outline_level(paragraph))
            self.assertNotIn("（约", text)
            self.assertTrue(text.endswith("段"), text)
            run = paragraph.runs[0]
            self.assertTrue(run.bold)
            self.assertEqual(run.font.size.pt, 12.0)
            self.assertEqual(run._element.rPr.rFonts.get(qn("w:eastAsia")), "楷体_GB2312")
            self.assertTrue(paragraph.paragraph_format.keep_with_next)


class T2hTableDisplayTests(unittest.TestCase):
    """T2h（表格显示效果）：tblLayout=fixed / 逐行 trHeight / 单元格 12pt 固定行距 /
    清单表类型列手动换行 / 高度②表数值不带 % / 标线②表里程取整。"""

    @staticmethod
    def _bundle() -> dict:
        marking = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                    "marking_position": "左侧标线", "value": 60.0, "target": 80.0, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        height = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                   "guardrail_type": "两波", "height": 600.0, "city": "测试市",
                   "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        marking.append({"category": "普通国省道", "route": "S120", "direction": "下行", "segment": "K6+000～K6+001",
                        "marking_position": "右侧标线", "value": 40.0, "target": 80.0, "city": "测试市",
                        "manager": "东莞市公路事务中心", "station_m": 6000.0, "end_m": 6001.0})
        return {"city": "测试市", "marking": marking, "height": height, "bolt": [], "notes": [],
                "comparison_detail": [], "weak_segments": [], "route_segments": []}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_tables_use_fixed_layout(self) -> None:
        """A1：每张表 tblLayout=fixed（缺它会按内容 auto-fit，窄列被压扁）。"""
        document = self._write()
        self.assertTrue(document.tables)
        for table in document.tables:
            node = table._tbl.find(qn("w:tblPr")).find(qn("w:tblLayout"))
            self.assertIsNotNone(node, table.rows[0].cells[0].text)
            self.assertEqual(node.get(qn("w:type")), "fixed")

    def test_every_row_has_template_height(self) -> None:
        """A2：每行都有 trHeight（atLeast；模板缺值 fallback=397），③ 表首行照模板 601。"""
        document = self._write()
        for table in document.tables:
            for row in table.rows:
                node = row._tr.find(qn("w:trPr")).find(qn("w:trHeight"))
                self.assertIsNotNone(node)
                self.assertGreater(int(node.get(qn("w:val"))), 0)
                self.assertEqual(node.get(qn("w:hRule")), "atLeast")
        captions = engine.gd_table_captions(document)
        for index, table in enumerate(document.tables):
            expected = engine.GD_TABLE_HEIGHTS_BY_CAPTION.get(captions[index] if index < len(captions) else "")
            if not expected:
                continue
            node = table.rows[0]._tr.find(qn("w:trPr")).find(qn("w:trHeight"))
            self.assertEqual(int(node.get(qn("w:val"))), expected[0], captions[index])
            data_heights = [int(r._tr.find(qn("w:trPr")).find(qn("w:trHeight")).get(qn("w:val")))
                            for r in table.rows[1:]]
            self.assertIn(expected[1], data_heights, captions[index])

    def test_cells_use_exact_line_spacing(self) -> None:
        """A3：单元格段落 12pt 固定行距（w:spacing line=240 lineRule=exact）。"""
        document = self._write()
        checked = 0
        for table in document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        node = paragraph._p.find(qn("w:pPr")).find(qn("w:spacing"))
                        self.assertIsNotNone(node)
                        self.assertEqual((node.get(qn("w:line")), node.get(qn("w:lineRule"))), ("240", "exact"))
                        checked += 1
        self.assertGreater(checked, 10)

    def test_list_type_column_breaks_line(self) -> None:
        """A4：清单表「类型」列照模板手动换行（高速\\n公路 / 普通\\n国省道）。"""
        document = self._write()
        table = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "类型")
        self.assertEqual({row.cells[0].text.strip() for row in table.rows[1:]},
                         {"高速\n公路", "普通\n国省道"})
        for row in table.rows[1:]:
            self.assertEqual(len(row.cells[0]._tc.findall(".//" + qn("w:br"))), 1)

    def test_height_table_has_no_percent_and_marking_km_is_integer(self) -> None:
        """B1/B3：高度②表数值不带 %（模板该表数据格即无 %）；标线②表里程列为整数。"""
        document = self._write()
        height_table = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "路线编号"
                            and any("三波合格率" in c.text for c in t.rows[0].cells))
        for row in height_table.rows[1:]:
            for cell in row.cells[1:]:
                self.assertNotIn("%", cell.text)
        marking_table = next(t for t in document.tables if t.rows[0].cells[0].text.strip() == "路线编号"
                             and any("左侧合格率" in c.text for c in t.rows[0].cells))
        for row in marking_table.rows[1:]:
            self.assertNotIn(".", row.cells[4].text)
            self.assertTrue(row.cells[4].text.strip().replace(",", "").isdigit(), row.cells[4].text)

    def test_section_margins_follow_template(self) -> None:
        """A5：版心照模板（L/R 31.75mm、T/B 25.40mm）。"""
        section = self._write().sections[0]
        self.assertEqual(section.left_margin.twips, 1800)
        self.assertEqual(section.right_margin.twips, 1800)
        self.assertEqual(section.top_margin.twips, 1440)
        self.assertEqual(section.bottom_margin.twips, 1440)

class T2iPercentHeightTests(unittest.TestCase):
    """T2i：全文档数据单元格不带 %（列头已写单位）；普通分支螺栓②表行高按模板别名条目取 (397, 432)。"""

    @staticmethod
    def _bundle() -> dict:
        marking = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                    "marking_position": "左侧标线", "value": 60.0, "target": 80.0, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0},
                   {"category": "普通国省道", "route": "S120", "direction": "下行", "segment": "K6+000～K6+001",
                    "marking_position": "右侧标线", "value": 40.0, "target": 80.0, "city": "测试市",
                    "manager": "东莞市公路事务中心", "station_m": 6000.0, "end_m": 6001.0}]
        height = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                   "guardrail_type": "两波", "height": 600.0, "city": "测试市",
                   "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        bolt = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                 "bolt_type": "拼接螺栓", "missing": 2, "expected": 100, "city": "测试市",
                 "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        return {"city": "测试市", "marking": marking, "height": height, "bolt": bolt, "notes": [],
                "comparison_detail": [], "weak_segments": [], "route_segments": []}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_no_percent_in_data_cells(self) -> None:
        """T2i/B1：列头含 % 的表，数据单元格一律不带 %（全文档统一，有意偏离模板混用）。"""
        document = self._write()
        offenders = [(table.rows[0].cells[0].text.strip(), cell.text.strip())
                     for table in document.tables for row in table.rows[1:]
                     for cell in row.cells if "%" in cell.text]
        self.assertEqual(offenders, [])

    def test_cell_formatters_have_no_percent(self) -> None:
        """T2i/B1：③/② 表格式化函数一律不带 %；空值分别写 `/` 与 `—`。"""
        writer = engine.GuangdongChapterWriter
        self.assertEqual(writer._g03_cell(None), "/")
        self.assertEqual(writer._g03_cell(0.1385), "13.85")
        self.assertEqual(writer._g03_cell(0.0), "0.00")
        self.assertEqual(writer._g03_num(None), "—")
        self.assertEqual(writer._g03_num(0.026, 2), "2.60")

    def test_bolt_route_table_heights_use_tsv_alias_entry(self) -> None:
        """T2i/A2：普通分支螺栓②表对应模板笔误同名的别名条目 (397, 432)；高速分支 510/510。"""
        self.assertEqual(engine.GD_TABLE_HEIGHTS_BY_CAPTION["普通国省道各路段公司螺栓缺失率汇总表"], (397, 432))
        self.assertEqual(engine.GD_TABLE_HEIGHTS_BY_CAPTION["高速公路各路段公司螺栓缺失率汇总表"], (510, 510))
        self.assertEqual(engine.GD_TABLE_HEIGHTS_BY_CAPTION["{市}交通安全设施抽检路段清单"], (397, 397))

class T2jStationSeparatorTests(unittest.TestCase):
    """T2j（用户第 5 轮 Q9）：桩号区间连接符统一 ASCII `-`，全文不得出现 `～`/`〜`/`~`。"""

    @staticmethod
    def _bundle() -> dict:
        marking = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                    "marking_position": "左侧标线", "value": 60.0, "target": 80.0, "city": "测试市",
                    "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        height = [{"category": "高速公路", "route": "S14", "direction": "上行", "segment": "K1+000～K1+001",
                   "guardrail_type": "两波", "height": 600.0, "city": "测试市",
                   "manager": "广东省高速公路有限公司", "station_m": 1000.0, "end_m": 1001.0}]
        return {"city": "测试市", "marking": marking, "height": height, "bolt": [], "notes": [],
                "comparison_detail": [], "weak_segments": [], "route_segments": []}

    def _write(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = engine.GuangdongChapterWriter.write(
                "测试市", self._bundle(), Path(temp_dir) / "out",
                Path(__file__).resolve().parents[1] / "templates" / "广东项目第五章模板.md",
                {"marking": 5, "height": 5, "bolt": 5},
            )
            return Document(output)

    def test_document_has_no_station_separator_variants(self) -> None:
        """T2j：正文与表格全文不得出现 `～`/`〜`/`~`（输入样例本身带 `～`）。"""
        document = self._write()
        texts = [p.text for p in document.paragraphs] + [
            cell.text for table in document.tables for row in table.rows for cell in row.cells]
        for variant in ("\uff5e", "\u301c", "\u007e"):
            self.assertEqual([t for t in texts if variant in t], [], repr(variant))
        self.assertTrue(any("-" in text for text in texts))

    def test_station_range_uses_ascii_hyphen(self) -> None:
        """T2j：② 表起止桩号用 `-`（`_gd03_range` 不再输出 `～`）。"""
        writer = engine.GuangdongChapterWriter
        self.assertEqual(writer._gd03_range(1000.0, 3000.0),
                         "%s-%s" % (engine.format_station(1000.0), engine.format_station(3000.0)))
        self.assertNotIn("～", writer._gd03_range(1000.0, 3000.0))
        self.assertEqual(writer._gd03_range(None, None), "—")

    def test_normalize_helper_replaces_variants_only(self) -> None:
        """T2j：helper 只把 `～`/`〜`/`~` 换成 `-`，`—`（占位/中文连接）保留。"""
        document = Document()
        document.add_paragraph("K1+000～K2+000、K3+000〜K4+000、K5+000~K6+000、路线—管养单位、无数据 —")
        changed = engine.normalize_station_separators(document)
        text = document.paragraphs[0].text
        self.assertEqual(changed, 1)
        self.assertEqual(text, "K1+000-K2+000、K3+000-K4+000、K5+000-K6+000、路线—管养单位、无数据 —")

class T2kWorkbookSeparatorTests(unittest.TestCase):
    """T2k：xlsx 写入点与 docx 共用同一归一化出口（桩号区间一律 ASCII `-`）。"""

    def test_text_normalizer_uses_hyphen(self) -> None:
        self.assertEqual(engine.normalize_station_text("K1+000～K2+000、K3+000~K4+000"),
                         "K1+000-K2+000、K3+000-K4+000")
        self.assertEqual(engine.normalize_station_text("无数据 —"), "无数据 —")

    def test_workbook_separators_normalized(self) -> None:
        from openpyxl import Workbook

        workbook = Workbook()
        sheet = workbook.active
        sheet["A1"], sheet["A2"], sheet["A3"] = "G0422下行K1005～K965", "K975+740~K979+740", "无数据 —"
        changed = engine.normalize_workbook_separators(workbook)
        self.assertEqual(changed, 2)
        self.assertEqual(sheet["A1"].value, "G0422下行K1005-K965")
        self.assertEqual(sheet["A2"].value, "K975+740-K979+740")
        self.assertEqual(sheet["A3"].value, "无数据 —")



class P1cTemplatePaletteAndRouteNameTests(unittest.TestCase):
    """P1c：① 统计图配色改用模板的 Office 2007-2010 调色板（brief-v3 需求 #4）；
    ② 叙述里路线号后补路线名称（需求 #3），③ 小标题不补。"""

    def test_series_palette_matches_template_theme(self) -> None:
        # 模板 18 个内嵌 xlsx 的 theme accent1/2/3（实测，见 reviews/p1c-narrative-charts.md）
        self.assertEqual(engine.GD_SERIES_COLORS, ("4F81BD", "C0504D", "9BBB59"))
        self.assertEqual(engine.GD_SERIES_HEX, ("#4F81BD", "#C0504D", "#9BBB59"))
        self.assertEqual(engine.GD03_COLORS, engine.GD_SERIES_HEX)
        # PIE_COLORS 前三位复用同一组；第 4 位 FFC000、第 5 位对齐 2007 系 accent5
        self.assertEqual(engine.PIE_COLORS[:3], list(engine.GD_SERIES_COLORS))
        self.assertEqual(engine.PIE_COLORS[3:], ["FFC000", "4BACC6"])
        # 模板根本不用的 Office 2013 色号，一个都不许残留在常量里
        for stale in ("4472C4", "ED7D31", "A5A5A5", "5B9BD5"):
            self.assertNotIn(stale, engine.GD_SERIES_COLORS)
            self.assertNotIn(stale, engine.PIE_COLORS)

    def test_no_office2013_palette_literal_left_in_engine_source(self) -> None:
        source = Path(engine.__file__).read_text(encoding="utf-8")
        body = "\n".join(line for line in source.splitlines()
                         if not line.lstrip().startswith("#"))   # 注释里保留取证说明
        for stale in ("4472C4", "ED7D31", "A5A5A5", "5B9BD5"):
            self.assertNotIn(stale, body, f"仍有 Office 2013 色号 {stale} 硬编码在引擎里")

    def test_intro_route_label_appends_authoritative_name(self) -> None:
        # 需求 #3：引言「涉及G0422、G1523…」必须写成「G0422武汉－深圳、G1523宁波－东莞」
        names = {"G0422": "武汉－深圳", "G1523": "宁波－东莞"}

        def label(route):
            return f"{route}{names.get(route) or ''}"

        self.assertEqual(label("G0422"), "G0422武汉－深圳")
        self.assertEqual("、".join(label(r) for r in ("G0422", "G1523")),
                         "G0422武汉－深圳、G1523宁波－东莞")
        # 查不到名称就只写路线号（如实标注为未命中），不拿别处措辞顶替
        self.assertEqual(label("G9999"), "G9999")

    def test_intro_km_split_matches_owner_triple(self) -> None:
        # P1k：引言「省集团X公里、非集团Y公里」必须与 ① 表三行同法取整（最大余数法），
        # 深圳曾出现 75+163=238 ≠ ① 表 237=163+74。
        self.assertEqual(engine.gd_km_split_text(237.283, 162.658, 74.625),
                         ("237", "163", "74"))
        self.assertEqual(sum(int(x) for x in engine.gd_km_split_text(237.283, 162.658, 74.625)[1:]), 237)
        # 普通分支无「全市=分项」三行关系，仍走原取整
        self.assertEqual(engine.gd_km_row_texts(
            [("XX市抽检普通国省道", {"km": 16.0}), ("XX市抽检普通国道", {"km": 8.0}),
             ("XX市抽检普通省道", {"km": 8.0})]), ["16", "8", "8"])
