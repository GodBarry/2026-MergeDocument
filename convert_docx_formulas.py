import zipfile
import sys
import os
import re
import argparse
import io
import concurrent.futures
import threading
from lxml import etree
from PIL import Image
import mathtype_converter
import subprocess
import tempfile

# Force UTF-8 encoding for subprocess to prevent GBK decoding errors on Windows
os.environ['PYTHONIOENCODING'] = 'utf-8'

def safe_subprocess_run(cmd, **kwargs):
    """
    Wrapper for subprocess.run that ensures UTF-8 encoding on Windows
    to prevent GBK decoding errors.
    """
    # Force UTF-8 encoding for stdout/stderr
    if 'encoding' not in kwargs:
        kwargs['encoding'] = 'utf-8'
    if 'errors' not in kwargs:
        kwargs['errors'] = 'replace'
    return subprocess.run(cmd, **kwargs)

# Optional imports for OCR
try:
    from pix2tex.cli import LatexOCR
    import latex2mathml.converter
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# Global lock for model inference if needed
model_lock = threading.Lock()

def convert_image_to_omml(image_data: bytes, model=None) -> str:
    """
    Convert image bytes (PNG/JPG/WMF) -> LaTeX -> MathML -> OMML
    """
    if not OCR_AVAILABLE:
        raise ImportError("pix2tex or latex2mathml not installed. OCR mode unavailable.")
    
    image = None
    try:
        with tempfile.NamedTemporaryFile(suffix=".tmp", delete=False) as tmp_in, \
             tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_out:
            
            tmp_in.write(image_data)
            tmp_in.flush()
            tmp_in.close()
            tmp_out.close()
            
            # Use ImageMagick to standardize input
            cmd = ["convert", "-density", "300", tmp_in.name, "-background", "white", "-alpha", "remove", "-alpha", "off", "-quality", "100", tmp_out.name]
            res = safe_subprocess_run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            if res.returncode == 0 and os.path.exists(tmp_out.name) and os.path.getsize(tmp_out.name) > 0:
                temp_img = Image.open(tmp_out.name)
                if temp_img.mode != "RGB":
                    temp_img = temp_img.convert("RGB")
                image = temp_img.copy()
            
            if os.path.exists(tmp_in.name): os.remove(tmp_in.name)
            if os.path.exists(tmp_out.name): os.remove(tmp_out.name)
    except Exception as e:
        pass # print(f"    [OCR Error] System conversion failed: {e}")

    if image is None: return None

    try:     
        if model is None: model = LatexOCR()
        latex = model(image)
        mathml = latex2mathml.converter.convert(latex)
        omml = mathtype_converter.convert_mathml_to_omml(mathml)
        return omml
    except Exception as e:
        pass # print(f"    [OCR Error] {e}")
        return None

def convert_image_task(args):
    rid, image_data, model = args
    try:
        omml = convert_image_to_omml(image_data, model)
        return rid, omml, None
    except Exception as e:
        return rid, None, str(e)

def convert_document(docx_path, strategy='binary', max_workers=None, scan_images=False, min_height_mm=2.5, max_height_mm=38, priority='binary', ocr_rescue=False):
    docx_path = os.path.abspath(docx_path)
    if not zipfile.is_zipfile(docx_path):
        print(f"Error: {docx_path} is not a valid docx.")
        return

    print(f"Processing {docx_path} with strategy: {strategy} (Priority: {priority}, Scan Images: {scan_images})")
    
    if (strategy in ['ocr', 'hybrid'] or scan_images) and not OCR_AVAILABLE:
         if strategy == 'ocr' or (strategy == 'hybrid' and priority == 'ocr'):
             print("Error: OCR dependencies not found.")
             return
         print("Warning: OCR unavailable. Hybrid mode will fallback to pure binary.")

    basename = os.path.basename(docx_path)
    output_path = os.path.join(os.path.dirname(docx_path), f"converted_{basename}")
    
    ocr_model = None
    if (strategy in ['ocr', 'hybrid'] or scan_images) and OCR_AVAILABLE:
        print("Loading OCR model...")
        try:
            ocr_model = LatexOCR()
        except Exception as e:
            print(f"Failed to load OCR model: {e}")
            if priority == 'ocr': return

    conversions = {} 
    rid_to_target = {}
    
    with zipfile.ZipFile(docx_path, 'r') as zf:
        if 'word/_rels/document.xml.rels' in zf.namelist():
            rels_content = zf.read('word/_rels/document.xml.rels').decode('utf-8')
            matches = re.findall(r'Id="(rId\d+)"[^>]*Target="([^"]+)"', rels_content)
            for rid, target in matches:
                target = 'word/' + target if not target.startswith('/') else target[1:]
                rid_to_target[rid] = target

        if 'word/document.xml' not in zf.namelist():
            print("Error: No document.xml found.")
            return
        xml_content = zf.read('word/document.xml')
        
    nsmap = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
        'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        'o': 'urn:schemas-microsoft-com:office:office',
        'v': 'urn:schemas-microsoft-com:vml',
        'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
        'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
    }
    
    tree = etree.fromstring(xml_content)
    ole_objects = tree.xpath(".//o:OLEObject", namespaces=nsmap)
    print(f"Found {len(ole_objects)} OLE objects phase 1.")
    
    ocr_tasks = [] 
    processed_rids = set()
    failed_rids = {}         # rid → (残缺 OMML, img_rid)
    failed_ocr = {}          # rid → OCR 补救结果
    count_success = 0
    count_binary = 0

    # Read binary data single-threaded first to avoid zip concurrency issues
    ole_tasks = []
    with zipfile.ZipFile(docx_path, 'r') as zf:
        for ole in ole_objects:
            rid = ole.get(f"{{{nsmap['r']}}}id")
            if not rid or rid not in rid_to_target: continue
            
            parent = ole.getparent()
            img_rid = None
            if parent is not None:
                 imagedata = parent.xpath(".//v:imagedata", namespaces=nsmap)
                 if imagedata:
                     img_rid = imagedata[0].get(f"{{{nsmap['r']}}}id")
            
            processed_rids.add(rid)
            if img_rid: processed_rids.add(img_rid)

            ole_path = rid_to_target[rid]
            try:
                ole_data = zf.read(ole_path) if ole_path.endswith('.bin') else None
                img_data = zf.read(rid_to_target[img_rid]) if img_rid and img_rid in rid_to_target else None
                ole_tasks.append((rid, img_rid, ole_path, ole_data, img_data))
            except: pass

    # Execute OLE conversions in parallel
    workers = max_workers or min(32, (os.cpu_count() or 1) + 4)
    print(f"Starting parallel OLE conversion with {workers} threads...")
    
    def convert_ole_worker(task):
        rid, img_rid, ole_path, ole_data, img_data = task
        omml = None
        is_binary_success = False
        is_failed_binary = False
        broken_omml = None
        needs_ocr_later = False
        needs_ocr_now = False

        methods = []
        if strategy == 'binary': methods = ['binary']
        elif strategy == 'ocr': methods = ['ocr']
        elif strategy == 'hybrid': methods = ['ocr', 'binary'] if priority == 'ocr' else ['binary', 'ocr']
                
        for method in methods:
            if omml: break
            if method == 'binary' and ole_data:
                try:
                    if mathtype_converter.is_mathtype_ole(ole_data):
                        res = mathtype_converter.convert_mathtype_to_omml(ole_data)
                        if res: 
                            if res.startswith('<?xml'): res = res.split('?>', 1)[1].strip()
                            if 'eqArr/>' in res or '<m:m/>' in res:
                                is_failed_binary = True
                                broken_omml = res
                            else:
                                omml = res
                                is_binary_success = True
                except: pass

            elif method == 'ocr':
                if strategy == 'hybrid' and priority == 'binary':
                    if img_data: needs_ocr_later = True
                    break
                else:
                    if img_data: needs_ocr_now = True
                    break
                    
        return rid, img_rid, omml, is_binary_success, is_failed_binary, broken_omml, needs_ocr_later, needs_ocr_now, img_data

    with concurrent.futures.ThreadPoolExecutor(max_workers=workers) as executor:
        for res_tuple in executor.map(convert_ole_worker, ole_tasks):
            rid, img_rid, omml, is_binary_success, is_failed_binary, broken_omml, needs_ocr_later, needs_ocr_now, img_data = res_tuple
            
            # handle needs_ocr_now synchronously to avoid model concurrency issues
            if needs_ocr_now and ocr_model and img_data:
                try:
                    ocr_res = convert_image_to_omml(img_data, ocr_model)
                    if ocr_res:
                        omml = ocr_res
                        is_binary_success = False
                        count_success += 1
                except: pass
                
            if omml: 
                conversions[rid] = omml
                if is_binary_success:
                    count_binary += 1
                sys.stdout.write(".")
            elif is_failed_binary:
                failed_rids[rid] = (broken_omml, img_rid)
                sys.stdout.write("!")
            elif needs_ocr_later:
                ocr_tasks.append((rid, img_data))
                
            sys.stdout.flush()

    if scan_images:
        print("\nScanning non-OLE images...")
        MAX_HEIGHT_EMU = max_height_mm * 36000   # 1mm = 36000 EMU
        MIN_HEIGHT_EMU = min_height_mm * 36000
        # 只找 w:drawing/wp:inline（段落中的行内图片，大概率是公式；跳过 wp:anchor 浮动大图）
        drawings = tree.xpath(".//w:drawing[wp:inline]", namespaces=nsmap)
        with zipfile.ZipFile(docx_path, 'r') as zf:
            for dwg in drawings:
                extents = dwg.xpath(".//wp:extent", namespaces=nsmap)
                if not extents: continue
                try:
                    cy = int(extents[0].get('cy', 0))
                    if not (MIN_HEIGHT_EMU < cy < MAX_HEIGHT_EMU): continue
                    
                    # 检查后面的段落是不是图标题（包含"图"或"Fig"），以跳过带有图题的连续结构
                    parent_p = dwg.xpath("ancestor::w:p", namespaces=nsmap)
                    is_figure_with_caption = False
                    if parent_p:
                        next_p = parent_p[0].getnext()
                        if next_p is not None and next_p.tag == f"{{{nsmap['w']}}}p":
                            text = "".join(next_p.itertext()).strip()
                            if text.startswith("图") or text.lower().startswith("fig"):
                                is_figure_with_caption = True
                    
                    if is_figure_with_caption:
                        continue

                    blips = dwg.xpath(".//a:blip", namespaces=nsmap)
                    if not blips: continue
                    img_rid = blips[0].get(f"{{{nsmap['r']}}}embed")
                    if img_rid and img_rid not in processed_rids and img_rid in rid_to_target:
                        ocr_tasks.append((img_rid, zf.read(rid_to_target[img_rid])))
                        processed_rids.add(img_rid)
                        sys.stdout.write("?")
                except: pass
        sys.stdout.flush()

    if ocr_tasks:
        print(f"\nStarting Parallel OCR for {len(ocr_tasks)} tasks...")
        args_list = [(rid, data, ocr_model) for (rid, data) in ocr_tasks]
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            for rid, res, err in executor.map(convert_image_task, args_list):
                if res:
                    conversions[rid] = res
                    count_success += 1
                    sys.stdout.write("o")
                else: sys.stdout.write("x")
                sys.stdout.flush()

    # ── OCR 补救：对二进制转换失败的公式尝试图片识别 ──────────────────
    if ocr_rescue and failed_rids and OCR_AVAILABLE:
        print(f"\n\n尝试对 {len(failed_rids)} 个失败公式进行 OCR 补救...")
        if ocr_model is None:
            try:
                ocr_model = LatexOCR()
            except Exception as e:
                print(f"  OCR 模型加载失败: {e}")
        if ocr_model:
            rescue_tasks = []
            with zipfile.ZipFile(docx_path, 'r') as zf:
                for frid, (broken_omml, f_img_rid) in failed_rids.items():
                    if f_img_rid and f_img_rid in rid_to_target:
                        try:
                            img_data = zf.read(rid_to_target[f_img_rid])
                            rescue_tasks.append((frid, img_data, ocr_model))
                        except Exception:
                            sys.stdout.write("X")
                            sys.stdout.flush()
            if rescue_tasks:
                with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
                    for frid, ocr_res, err in executor.map(convert_image_task, rescue_tasks):
                        if ocr_res:
                            failed_ocr[frid] = ocr_res
                            sys.stdout.write("O")  # OCR 成功
                        else:
                            sys.stdout.write("X")  # OCR 也失败
                        sys.stdout.flush()
    elif ocr_rescue and failed_rids and not OCR_AVAILABLE:
        print(f"\n  OCR 不可用，跳过补救（安装 pix2tex 可启用）")
    elif not ocr_rescue and failed_rids:
        print(f"\n  跳过 OCR 补救（使用 --ocr-rescue 可启用）")

    print(f"\n\n转换统计: 成功 {count_binary + count_success} (二进制:{count_binary}, OCR:{count_success}), 失败(空矩阵):{len(failed_rids)}, OCR补救:{len(failed_ocr)}")

    if not conversions and not failed_rids:
        print("No conversions made.")
        return

    W = nsmap['w']

    def _make_marker_para(label_text, omml_str=None, color_hex='FF0000'):
        """创建一个彩色标注段落: [标注文字] + (可选的公式)"""
        p = etree.Element(f'{{{W}}}p')
        # 标注文字 run
        r = etree.SubElement(p, f'{{{W}}}r')
        rPr = etree.SubElement(r, f'{{{W}}}rPr')
        color = etree.SubElement(rPr, f'{{{W}}}color')
        color.set(f'{{{W}}}val', color_hex)
        etree.SubElement(rPr, f'{{{W}}}b')
        sz = etree.SubElement(rPr, f'{{{W}}}sz')
        sz.set(f'{{{W}}}val', '18')  # 9pt
        t = etree.SubElement(r, f'{{{W}}}t')
        t.text = label_text

        # 插入 OMML 公式（着色）
        if omml_str:
            try:
                if omml_str.startswith('<?xml'):
                    omml_str = omml_str.split('?>', 1)[1].strip()
                omml_el = etree.fromstring(omml_str)
                M = nsmap['m']
                for mr in omml_el.iter(f'{{{M}}}r'):
                    mrPr = mr.find(f'{{{M}}}rPr')
                    if mrPr is None:
                        mrPr = etree.SubElement(mr, f'{{{M}}}rPr')
                        mr.insert(0, mrPr)
                    wrPr = mrPr.find(f'{{{W}}}rPr')
                    if wrPr is None:
                        wrPr = etree.SubElement(mrPr, f'{{{W}}}rPr')
                    wcolor = wrPr.find(f'{{{W}}}color')
                    if wcolor is None:
                        wcolor = etree.SubElement(wrPr, f'{{{W}}}color')
                    wcolor.set(f'{{{W}}}val', color_hex)
                p.append(omml_el)
            except Exception:
                pass
        return p

    print("Patching document...")
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    doc_tree = etree.fromstring(zin.read(item.filename))

                    # 1. 替换成功转换的公式
                    for rid, omml_str in conversions.items():
                        if omml_str.startswith('<?xml'): omml_str = omml_str.split('?>', 1)[1]
                        for ole in doc_tree.xpath(f".//o:OLEObject[@r:id='{rid}']", namespaces=nsmap):
                            target = ole
                            parent = ole.getparent()
                            if parent is not None and parent.tag.endswith('}object'): target = parent
                            try:
                                target.getparent().replace(target, etree.fromstring(omml_str))
                            except: pass
                        for dwg in doc_tree.xpath(f".//w:drawing[descendant::a:blip/@r:embed='{rid}']", namespaces=nsmap):
                            try:
                                dwg.getparent().replace(dwg, etree.fromstring(omml_str))
                            except: pass

                    # 2. 对失败公式：保留原始 OLE，插入标注行
                    for frid, (broken_omml, _) in failed_rids.items():
                        for ole in doc_tree.xpath(f".//o:OLEObject[@r:id='{frid}']", namespaces=nsmap):
                            node = ole
                            para = None
                            while node is not None:
                                if node.tag == f'{{{W}}}p':
                                    para = node
                                    break
                                node = node.getparent()
                            if para is not None:
                                parent_of_para = para.getparent()
                                if parent_of_para is not None:
                                    idx = list(parent_of_para).index(para)
                                    insert_at = idx + 1
                                    # 红色：二进制转换残缺结果
                                    m1 = _make_marker_para(
                                        '【⚠ 二进制转换结果（残缺）：】',
                                        broken_omml, 'FF0000'
                                    )
                                    parent_of_para.insert(insert_at, m1)
                                    insert_at += 1
                                    # 蓝色：OCR 补救结果（如果有）
                                    if frid in failed_ocr:
                                        m2 = _make_marker_para(
                                            '【🔍 OCR 识别结果：】',
                                            failed_ocr[frid], '0070C0'
                                        )
                                        parent_of_para.insert(insert_at, m2)

                    zout.writestr(item, etree.tostring(doc_tree, encoding='utf-8', xml_declaration=True))
                else:
                    zout.writestr(item, zin.read(item.filename))
    print(f"Created {output_path}")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Convert MathType to OMML in DOCX")
    parser.add_argument("--docx_file", default='BY1617126.docx', help="Path to docx file")
    parser.add_argument("--strategy", choices=['binary', 'ocr', 'hybrid'], default='binary')
    parser.add_argument("--priority", choices=['binary', 'ocr'], default='binary')
    parser.add_argument("--workers", type=int, default=None)
    parser.add_argument("--scan-images", action="store_true")
    parser.add_argument("--ocr-rescue", action="store_true", default=False,
                        help="对二进制转换失败的公式尝试 OCR 补救（默认关闭）")
    parser.add_argument("--min-height", type=float, default=2.5, help="最小图片高度(mm)")
    parser.add_argument("--max-height", type=float, default=38, help="最大图片高度(mm)")
    args = parser.parse_args()

    convert_document(args.docx_file, args.strategy, args.workers, args.scan_images,
                     args.min_height, args.max_height, args.priority, args.ocr_rescue)
