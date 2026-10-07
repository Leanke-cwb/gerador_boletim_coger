from pathlib import Path
import os, sys
import cv2
import numpy as np
import pytesseract

SISTEMA = [r"C:\Program Files\Tesseract-OCR\tesseract.exe", r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe"]

def _runtime_base():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return Path(__file__).resolve().parent.parent

def _set_tessdata(folder):
    td = Path(folder) / "tessdata"
    if td.exists():
        os.environ["TESSDATA_PREFIX"] = str(td)

def localizar_tesseract():
    base = _runtime_base()
    for p in [base / "tesseract" / "tesseract.exe", base / "vendor" / "tesseract" / "tesseract.exe"]:
        if p.exists():
            pytesseract.pytesseract.tesseract_cmd = str(p)
            _set_tessdata(p.parent)
            return str(p)
    atual = getattr(pytesseract.pytesseract, "tesseract_cmd", None)
    if atual and atual != "tesseract" and Path(atual).exists():
        _set_tessdata(Path(atual).parent)
        return atual
    for s in SISTEMA:
        p = Path(s)
        if p.exists():
            pytesseract.pytesseract.tesseract_cmd = str(p)
            _set_tessdata(p.parent)
            return str(p)
    try:
        pytesseract.pytesseract.tesseract_cmd = "tesseract"
        pytesseract.get_tesseract_version()
        return "tesseract (PATH)"
    except Exception:
        return None

def _configurar_tesseract():
    if not localizar_tesseract():
        raise FileNotFoundError("Tesseract OCR não encontrado. Reinstale o Gerador de Boletim.")

def _preprocessar(caminho):
    img = cv2.imread(str(caminho))
    if img is None:
        raise ValueError("Não foi possível abrir a imagem.")
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3, 3), 0)
    bw = cv2.adaptiveThreshold(gray, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 15)
    return img, gray, bw

def _detectar_grade(bw):
    inv = 255 - bw
    hk = cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, bw.shape[1] // 25), 1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(20, bw.shape[0] // 25)))
    h = cv2.morphologyEx(inv, cv2.MORPH_OPEN, hk, iterations=1)
    v = cv2.morphologyEx(inv, cv2.MORPH_OPEN, vk, iterations=1)
    grade = cv2.add(h, v)
    return grade, float(np.count_nonzero(grade)) / float(grade.size)

def _ocr_texto(gray):
    return "\n".join(l.rstrip() for l in pytesseract.image_to_string(gray, lang="por", config="--oem 3 --psm 6").splitlines()).strip()

def _agrupar(celulas, tolerancia=12):
    grupos=[]
    for cel in sorted(celulas, key=lambda c:(c[1],c[0])):
        x,y,w,h,texto=cel; cy=y+h//2
        for g in grupos:
            if abs(cy-g['y'])<=tolerancia:
                g['cells'].append(cel); g['y']=(g['y']+cy)//2; break
        else:
            grupos.append({'y':cy,'cells':[cel]})
    out=[]
    for g in sorted(grupos,key=lambda g:g['y']):
        row=[c[4] for c in sorted(g['cells'],key=lambda c:c[0])]
        if any(v.strip() for v in row): out.append(row)
    return out

def _ocr_tabela(img,bw,grade):
    contornos,_=cv2.findContours(grade, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    hi,wi=bw.shape[:2]; cel=[]
    for cnt in contornos:
        x,y,w,h=cv2.boundingRect(cnt)
        if (w>wi*.95 and h>hi*.80) or w<35 or h<18 or w>wi*.95 or h>hi*.50: continue
        x1,y1,x2,y2=max(0,x+3),max(0,y+3),min(wi,x+w-3),min(hi,y+h-3)
        if x2<=x1 or y2<=y1: continue
        roi=cv2.cvtColor(img[y1:y2,x1:x2],cv2.COLOR_BGR2GRAY)
        t=" ".join(pytesseract.image_to_string(roi,lang="por",config="--oem 3 --psm 6").strip().split())
        cel.append((x,y,w,h,t))
    linhas=_agrupar(cel); seen=set(); out=[]
    for r in linhas:
        k=tuple(r)
        if k not in seen: out.append(r); seen.add(k)
    return out

def processar_imagem(caminho):
    _configurar_tesseract()
    img,gray,bw=_preprocessar(caminho)
    grade,ratio=_detectar_grade(bw)
    if ratio>0.008:
        tabela=_ocr_tabela(img,bw,grade)
        if len(tabela)>=2 and max((len(r) for r in tabela),default=0)>=2:
            return {"tipo":"tabela","tabela":tabela,"confianca_layout":"aproximada"}
    return {"tipo":"texto","texto":_ocr_texto(gray),"confianca_layout":"texto"}
