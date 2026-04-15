# -*- coding: utf-8 -*-
"""
excel_export.py — Excel 生成服务
职责：xlsxwriter 主方案 + COM 回退，不包含 HTTP 或数据存取逻辑
"""

import io
import logging
import shutil
import zipfile
import re
import os
import tempfile
from datetime import datetime
from pathlib import Path

logger = logging.getLogger(__name__)

# ---- 依赖检查 ----
try:
    import xlsxwriter
    HAS_XLSXWRITER = True
except ImportError:
    HAS_XLSXWRITER = False
    logger.warning("缺少 xlsxwriter，请运行: pip install xlsxwriter")

try:
    from PIL import Image as PILImage
    HAS_PIL = True
except ImportError:
    HAS_PIL = False
    logger.error("缺少 Pillow，请运行: pip install Pillow")

# ---- Office COM 检测 ----
HAS_WIN32COM = False
OFFICE_TYPE = None   # 'excel' | 'wps' | None


def detect_office_tool():
    """检测可用的 Office COM 工具（Excel 优先，其次 WPS）"""
    global HAS_WIN32COM, OFFICE_TYPE

    candidates = [
        ('Excel.Application', 'excel'),
        ('ET.Application',    'wps'),
        ('WPS.Application',   'wps'),
    ]
    for prog_id, kind in candidates:
        try:
            import win32com.client
            app = win32com.client.Dispatch(prog_id)
            app.Quit()
            HAS_WIN32COM = True
            OFFICE_TYPE = kind
            logger.info(f"检测到 Office COM: {prog_id}")
            return True
        except Exception:
            pass

    HAS_WIN32COM = False
    OFFICE_TYPE = None
    logger.info("未检测到 Office COM，将使用 xlsxwriter")
    return False


def get_excel_creator_info():
    """返回当前 Excel 生成器信息"""
    if HAS_XLSXWRITER:
        return {'type': 'xlsxwriter', 'tool': 'xlsxwriter'}
    if HAS_WIN32COM:
        return {'type': 'com', 'tool': OFFICE_TYPE}
    return {'type': 'none', 'tool': 'none'}


# ---------------------------------------------------------------------------
# xlsxwriter 方案
# ---------------------------------------------------------------------------

def _write_sheet_xlsx(wb, ws, records, img_map):
    """用 xlsxwriter 写单个 sheet"""
    header_fmt = wb.add_format({
        'bold': True, 'font_size': 12, 'font_name': '微软雅黑',
        'font_color': '#000000', 'bg_color': '#FFFFFF',
        'align': 'center', 'valign': 'vcenter', 'border': 1,
    })
    data_fmt = wb.add_format({
        'font_size': 10, 'font_name': '微软雅黑',
        'font_color': '#000000', 'align': 'center', 'valign': 'vcenter', 'border': 1,
    })
    money_fmt = wb.add_format({
        'font_size': 10, 'font_name': '微软雅黑',
        'font_color': '#000000', 'align': 'center', 'valign': 'vcenter',
        'num_format': '¥#,##0.00', 'border': 1,
    })
    total_label_fmt = wb.add_format({
        'bold': True, 'font_size': 10, 'font_name': '微软雅黑',
        'font_color': '#000000', 'align': 'right', 'valign': 'vcenter', 'border': 1,
    })
    total_money_fmt = wb.add_format({
        'bold': True, 'font_size': 10, 'font_name': '微软雅黑',
        'font_color': '#000000', 'align': 'center', 'valign': 'vcenter',
        'num_format': '¥#,##0.00', 'border': 1,
    })

    # 列宽
    for col, w in enumerate([12, 16, 14, 18, 12, 15, 8, 14]):
        ws.set_column(col, col, w)

    # 表头
    headers = ['时间', '产品', '关联项目', '发生原因', '金额', '详情截图', '是否有票', '开票主体']
    for col, h in enumerate(headers):
        ws.write(0, col, h, header_fmt)
    ws.set_row(0, 40)

    # 数据行
    total = 0.0
    for idx, r in enumerate(records):
        excel_row = idx + 1
        data_row  = excel_row + 1   # 1-based 行号（header=1，数据从2开始）
        total += r.get('amount', 0)

        ws.write(excel_row, 0, r.get('time', ''),                data_fmt)
        ws.write(excel_row, 1, r.get('product', ''),             data_fmt)
        ws.write(excel_row, 2, r.get('related', '') or '-',      data_fmt)
        ws.write(excel_row, 3, r.get('reason', ''),              data_fmt)
        ws.write_number(excel_row, 4, r.get('amount', 0),        money_fmt)
        ws.write(excel_row, 6, r.get('hasTicket', ''),           data_fmt)
        ws.write(excel_row, 7, r.get('ticketEntity', '') or '-', data_fmt)
        ws.set_row(excel_row, 200)

        if data_row in img_map:
            img_path, pil_img = img_map[data_row]
            orig_w, orig_h = pil_img.size
            scale = (200.0 / orig_h) if orig_h > 0 else 1.0
            ws.insert_image(excel_row, 5, img_path, {
                'x_scale': scale, 'y_scale': scale, 'object_position': 2,
            })
        else:
            ws.write(excel_row, 5, '', data_fmt)

    # 合计行
    if records:
        total_row = len(records) + 1
        ws.write(total_row, 3, '合计：', total_label_fmt)
        ws.write_number(total_row, 4, total, total_money_fmt)
        ws.set_row(total_row, 40)


def _prepare_images_xlsx(records, prefix, img_tmp_dir):
    """解码 base64 图片并保存到临时目录，返回 {excel_row: (img_path, pil_img)}"""
    img_map = {}
    for idx, r in enumerate(records):
        if not r.get('image'):
            continue
        try:
            import base64
            data = r['image']
            if ',' in data:
                _, data = data.split(',', 1)
            img_bytes = base64.b64decode(data)
            pil_img = PILImage.open(io.BytesIO(img_bytes))
            img_path = img_tmp_dir / f'{prefix}_{idx}.png'
            pil_img.save(str(img_path), format='PNG')
            img_map[idx + 2] = (str(img_path), pil_img)
        except Exception as e:
            logger.error(f"准备图片失败 (record {idx}): {e}")
    return img_map


def create_excel_xlsxwriter(data, output_dir: Path):
    """xlsxwriter 方案生成 Excel，返回 (out_path, filename)"""
    if not HAS_XLSXWRITER:
        raise RuntimeError("缺少 xlsxwriter，请运行: pip install xlsxwriter")
    if not HAS_PIL:
        raise RuntimeError("缺少 Pillow，请运行: pip install Pillow")

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename  = f'报销汇总_{timestamp}.xlsx'
    out_path  = output_dir / filename

    expense_records   = data.get('expense', [])
    reimburse_records = data.get('reimburse', [])

    img_tmp_dir = output_dir / '_img_tmp'
    img_tmp_dir.mkdir(exist_ok=True)

    try:
        expense_imgs   = _prepare_images_xlsx(expense_records,   'expense',   img_tmp_dir)
        reimburse_imgs = _prepare_images_xlsx(reimburse_records, 'reimburse', img_tmp_dir)

        wb = xlsxwriter.Workbook(str(out_path))

        if expense_records:
            _write_sheet_xlsx(wb, wb.add_worksheet('费用模板'), expense_records, expense_imgs)
        if reimburse_records:
            _write_sheet_xlsx(wb, wb.add_worksheet('报销模板'), reimburse_records, reimburse_imgs)
        if not expense_records and not reimburse_records:
            wb.add_worksheet('费用模板').write(0, 0, '无数据')

        wb.close()
    finally:
        shutil.rmtree(img_tmp_dir, ignore_errors=True)

    logger.info(f"xlsxwriter 生成 Excel 成功: {out_path}")
    return out_path, filename


# ---------------------------------------------------------------------------
# COM 回退方案
# ---------------------------------------------------------------------------

def create_excel_com(data, output_dir: Path):
    """COM 方案生成 Excel（需要 Excel 或 WPS 已安装），返回 (out_path, filename)"""
    if not HAS_WIN32COM:
        raise RuntimeError("未检测到 Office COM 接口")
    if not HAS_PIL:
        raise RuntimeError("缺少 Pillow 库")

    import win32com.client
    import base64

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename  = f'报销汇总_{timestamp}.xlsx'
    out_path  = output_dir / filename

    expense_records   = data.get('expense', [])
    reimburse_records = data.get('reimburse', [])

    img_tmp_dir = output_dir / '_img_tmp'
    img_tmp_dir.mkdir(exist_ok=True)

    def _save_images(records, prefix):
        img_map = {}
        for idx, r in enumerate(records):
            if not r.get('image'):
                continue
            try:
                raw = r['image']
                if ',' in raw:
                    _, raw = raw.split(',', 1)
                pil_img = PILImage.open(io.BytesIO(base64.b64decode(raw)))
                p = img_tmp_dir / f'{prefix}_{idx}.png'
                pil_img.save(str(p), format='PNG')
                img_map[idx + 2] = (str(p), pil_img)
            except Exception as e:
                logger.error(f"COM 准备图片失败 (record {idx}): {e}")
        return img_map

    expense_imgs   = _save_images(expense_records,   'expense')
    reimburse_imgs = _save_images(reimburse_records, 'reimburse')

    prog_id = 'Excel.Application' if OFFICE_TYPE == 'excel' else 'ET.Application'
    app = None
    try:
        app = win32com.client.Dispatch(prog_id)
        app.Visible = False
        app.DisplayAlerts = False
        wb = app.Workbooks.Add()

        while wb.Sheets.Count > 1:
            wb.Sheets(wb.Sheets.Count).Delete()

        headers = ['时间', '产品', '关联项目', '发生原因', '金额', '详情截图', '是否有票', '开票主体']

        def _write_sheet_com(ws, sheet_name, records, img_map):
            ws.Name = sheet_name
            for col_idx, h in enumerate(headers, 1):
                c = ws.Cells(1, col_idx)
                c.Value = h
                c.Font.Bold = True
                c.Font.Size = 12
                c.Font.Name = '微软雅黑'
                c.Interior.ColorIndex = -4142
                c.Font.Color = 0x000000
                c.HorizontalAlignment = -4108
                c.VerticalAlignment   = -4160

            for col_idx, w in enumerate([12, 16, 14, 18, 12, 22, 8, 14], 1):
                try:
                    ws.Columns(col_idx).ColumnWidth = w
                except Exception as e:
                    logger.warning(f"设置列宽失败 col={col_idx}: {e}")

            try:
                ws.Rows(1).RowHeight = 40
            except Exception as e:
                logger.warning(f"设置表头行高失败: {e}")

            total = 0.0
            for row_idx, r in enumerate(records, 2):
                total += r.get('amount', 0)

                def sc(cell):
                    cell.HorizontalAlignment = -4108
                    cell.VerticalAlignment   = -4160
                    cell.Font.Size = 10
                    cell.Font.Name = '微软雅黑'

                for col, key, default in [
                    (1, 'time', ''), (2, 'product', ''), (3, 'related', '-'),
                    (4, 'reason', ''), (7, 'hasTicket', ''), (8, 'ticketEntity', '-'),
                ]:
                    c = ws.Cells(row_idx, col)
                    c.Value = r.get(key, '') or default
                    sc(c)

                c5 = ws.Cells(row_idx, 5)
                c5.Value = r.get('amount', 0)
                c5.NumberFormat = '¥#,##0.00'
                sc(c5)

                try:
                    ws.Rows(row_idx).RowHeight = 300
                except Exception as e:
                    logger.warning(f"设置行高失败 row={row_idx}: {e}")

                if row_idx in img_map:
                    try:
                        img_path, pil_img = img_map[row_idx]
                        orig_w, orig_h = pil_img.size
                        cell = ws.Cells(row_idx, 6)
                        cw, ch = cell.Width, cell.Height
                        if orig_w > 0 and orig_h > 0:
                            scale = min(cw / orig_w, ch / orig_h, 1.0)
                            dw, dh = orig_w * scale, orig_h * scale
                        else:
                            dw, dh = cw - 4, ch - 4
                        pic = ws.Shapes.AddPicture(
                            img_path, LinkToFile=False, SaveWithDocument=True,
                            Left=cell.Left + (cw - dw) / 2,
                            Top=cell.Top  + (ch - dh) / 2,
                            Width=dw, Height=dh,
                        )
                        pic.Placement = 1
                    except Exception as e:
                        logger.error(f"COM 插入图片失败 row={row_idx}: {e}")
                        ws.Cells(row_idx, 6).Value = '图片加载失败'

            total_row = len(records) + 2
            for col, val, fmt in [(4, '合计：', None), (5, total, '¥#,##0.00')]:
                c = ws.Cells(total_row, col)
                c.Value = val
                c.Font.Bold = True
                c.Font.Size = 10
                c.Font.Name = '微软雅黑'
                c.HorizontalAlignment = -4108
                c.VerticalAlignment   = -4160
                if fmt:
                    c.NumberFormat = fmt
            try:
                ws.Rows(total_row).RowHeight = 300
            except Exception:
                pass

        if expense_records:
            _write_sheet_com(wb.Worksheets(1), '费用模板', expense_records, expense_imgs)
        if reimburse_records:
            ws2 = wb.Worksheets.Add() if wb.Sheets.Count < 2 else wb.Worksheets(2)
            _write_sheet_com(ws2, '报销模板', reimburse_records, reimburse_imgs)
        if not expense_records and not reimburse_records:
            wb.Worksheets(1).Name = '费用模板'
            wb.Worksheets(1).Cells(1, 1).Value = '无数据'

        wb.SaveAs(str(out_path), 51)
        wb.Close()
    finally:
        if app:
            try:
                app.Quit()
            except Exception:
                pass
        shutil.rmtree(img_tmp_dir, ignore_errors=True)

    logger.info(f"COM 生成 Excel 成功: {out_path}")
    return out_path, filename


# ---------------------------------------------------------------------------
# 统一入口
# ---------------------------------------------------------------------------

def create_excel(data, output_dir: Path):
    """统一 Excel 生成入口：xlsxwriter 优先，COM 回退"""
    if HAS_XLSXWRITER:
        try:
            return create_excel_xlsxwriter(data, output_dir)
        except Exception as e:
            logger.warning(f"xlsxwriter 失败，回退 COM: {e}")
            if HAS_WIN32COM:
                return create_excel_com(data, output_dir)
            raise
    if HAS_WIN32COM:
        return create_excel_com(data, output_dir)
    raise RuntimeError("无可用 Excel 生成引擎（请安装 xlsxwriter 或 Office/WPS）")


# ---------------------------------------------------------------------------
# XML 后处理：twoCellAnchor → oneCellAnchor（备用，当前不启用）
# ---------------------------------------------------------------------------

def embed_images_into_cells(xlsx_path):
    """将 xlsxwriter 生成的浮动图片改为单元格锚定（实验性）"""
    tmp_dir = tempfile.mkdtemp(prefix='xlsx_embed_')
    try:
        with zipfile.ZipFile(xlsx_path, 'r') as z:
            z.extractall(tmp_dir)

        drawings_dir = os.path.join(tmp_dir, 'xl', 'drawings')
        if not os.path.isdir(drawings_dir):
            return

        for fname in os.listdir(drawings_dir):
            if not fname.endswith('.xml'):
                continue
            fpath = os.path.join(drawings_dir, fname)
            with open(fpath, 'r', encoding='utf-8') as f:
                xml = f.read()
            orig = xml
            xml = re.sub(r'<xdr:twoCellAnchor[^>]*>', '<xdr:oneCellAnchor>', xml)
            xml = xml.replace('</xdr:twoCellAnchor>', '</xdr:oneCellAnchor>')
            xml = re.sub(r'<xdr:to>.*?</xdr:to>', '', xml, flags=re.DOTALL)
            if xml != orig:
                with open(fpath, 'w', encoding='utf-8') as f:
                    f.write(xml)
                logger.info(f"图片已改为单元格嵌入: {fname}")

        tmp_xlsx = str(xlsx_path) + '.tmp'
        with zipfile.ZipFile(tmp_xlsx, 'w', compression=zipfile.ZIP_DEFLATED) as z:
            for root, _, files in os.walk(tmp_dir):
                for file in files:
                    fp = os.path.join(root, file)
                    z.write(fp, os.path.relpath(fp, tmp_dir))
        os.replace(tmp_xlsx, xlsx_path)
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
