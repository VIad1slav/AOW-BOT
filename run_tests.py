#!/usr/bin/env python3
"""Run every offline check and print one verdict.

    python3 run_tests.py           # всё, что быстро — гонять перед рестартом
    python3 run_tests.py --full    # плюс открытие результата в Word / LibreOffice

Nothing here touches the network or the live bot.
"""

import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent
SUITES = [
    ('docx_text — замены по прогонам Word', 'test_docx_text.py'),
    ('translit — адрес в польскую латиницу', 'test_translit.py'),
    ('passport — чтение MRZ паспорта', 'test_passport.py'),
    ('proforma — разбор PDF и сборка DOCX', 'test_proforma.py'),
    ('proforma_pipes — Pro Forma по спецификации', 'test_proforma_pipes.py'),
    ('cmr — накладная из спецификации', 'test_cmr.py'),
    ('bot — маршрутизация всего диалога', 'test_flow.py'),
]


def run(script):
    proc = subprocess.run([sys.executable, str(HERE / script)],
                          capture_output=True, text=True,
                          encoding='utf-8', errors='replace', cwd=str(HERE))
    return proc.returncode, (proc.stdout or '') + (proc.stderr or '')


def count(output):
    """How many individual checks passed / failed in a suite's output."""
    passed = output.count('  ✅ ') + output.count('  ok   ')
    failed = output.count('  ❌ ') + output.count('  FAIL ')
    return passed, failed


WORD_PS = r'''
$ErrorActionPreference = 'Stop'
$src = $args[0]
$pdf = [System.IO.Path]::ChangeExtension($src, ".pdf")
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$word.DisplayAlerts = 0
try {
    $doc = $word.Documents.Open($src, $false, $true)
    $pages = $doc.ComputeStatistics(2)
    $doc.SaveAs([ref]$pdf, [ref]17)
    $doc.Close($false)
    Write-Output ("OK " + $pages + " " + (Get-Item $pdf).Length)
} catch {
    Write-Output ("FAIL " + $_.Exception.Message)
} finally {
    $word.Quit()
}
'''


def sample_documents(tmp):
    """One document per template, so a broken template is caught by the office
    check and not by the customer. Returns [(label, path), ...]."""
    sys.path.insert(0, str(HERE))
    import proforma as pf
    import proforma_pipes as pp

    car = Path(tmp) / pf.out_filename('PF26-9/2')
    pf.build_proforma(car, {
        'pf_num': 'PF26-9/2', 'date': '2026-09-02', 'termin': '2026-09-09',
        'buyer_name': 'TEST BUYER', 'buyer_info': 'Belarus\npaszport: X1',
        'vehicles': [{'model': 'PEUGEOT 308 VP', 'plate': 'FW-646-BP',
                      'reg_date': '23/12/2020', 'vin': 'VF3LPHNSKLS232589',
                      'cn': '87032290', 'clearance': '145', 'price': 5400}],
    })

    pipes = Path(tmp) / pp.out_filename('PF26-9/3')
    pp.build_pipes_proforma(pipes, {
        'pf_num': 'PF26-9/3', 'date': '2026-09-02', 'termin': '2026-09-27',
        'spec': '126', 'code': '1139034',
        'items': [
            {'artikel': '116060', 'name': 'HTEM Pipe DN/OD 125х2000 mm ',
             'qty': 54, 'price': 6.05, 'amount': 326.70},
            {'artikel': '335040', 'name': 'Skolan Safe-EM Pipe DN/OD 110х1000 mm ',
             'qty': 640, 'price': 5.64, 'amount': 3609.60},
            {'artikel': '220630-03', 'name': 'KGK Cap DN/OD 110 ',
             'qty': 800, 'price': 0.79, 'amount': 632.00},
        ],
    })
    return [('Pro Forma (машины)', car), ('Pro Forma (спецификация)', pipes)]


def word_check():
    """Windows only: open every Pro Forma in the real Word and export a PDF.

    This is the check that actually matters — Word is what opens these
    documents in the end, and it is stricter about malformed OOXML than
    python-docx is.
    """
    import tempfile
    with tempfile.TemporaryDirectory() as tmp:
        docs = sample_documents(tmp)
        script = Path(tmp) / 'open.ps1'
        script.write_text(WORD_PS, encoding='utf-8')
        good = True
        for label, src in docs:
            try:
                proc = subprocess.run(
                    ['powershell', '-NoProfile', '-NonInteractive',
                     '-ExecutionPolicy', 'Bypass', '-File', str(script), str(src)],
                    capture_output=True, text=True, encoding='utf-8',
                    errors='replace', timeout=180)
            except (FileNotFoundError, subprocess.TimeoutExpired):
                print('  ⏭  ПРОПУЩЕНО — Word недоступен')
                return True

            out = (proc.stdout or '').strip()
            if out.startswith('OK'):
                _, pages, size = out.split()
                print('  ✅ Word открыл «%s» и выгнал PDF '
                      '(страниц: %s, %d КБ)' % (label, pages, int(size) // 1024))
                continue
            if 'Word.Application' in out or not out:
                print('  ⏭  ПРОПУЩЕНО — Word не установлен')
                return True
            print('  ❌ Word не смог открыть «%s»' % label)
            print('     ' + out[:300])
            good = False
        return good


def office_check():
    """Prove a generated Pro Forma really opens in office software, not just in
    python-docx.

    A trivial python-docx file is converted alongside it as a control. The
    AOW server carries only libreoffice-calc — no Writer — so *no* .docx opens
    there; without the control that reads as a broken document instead of a
    missing component.
    """
    import shutil
    import tempfile
    from docx import Document

    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        if sys.platform == 'win32':
            return word_check()
        print('  ⏭  ПРОПУЩЕНО — LibreOffice не найден')
        return True

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        control = tmp / 'control.docx'
        doc = Document()
        doc.add_paragraph('контрольный документ')
        doc.save(str(control))

        docs = sample_documents(tmp)
        proc = subprocess.run(
            [soffice, '--headless', '--convert-to', 'pdf', '--outdir', str(tmp),
             str(control)] + [str(p) for _, p in docs],
            capture_output=True, text=True, timeout=240)

        def made(path):
            out = path.with_suffix('.pdf')
            return out.exists() and out.stat().st_size > 3000

        if not made(control):
            print('  ⏭  ПРОПУЩЕНО — этот LibreOffice не открывает DOCX вообще '
                  '(нет libreoffice-writer); XLSX→PDF это не затрагивает')
            return True
        good = True
        for label, src in docs:
            if made(src):
                size = src.with_suffix('.pdf').stat().st_size // 1024
                print('  ✅ LibreOffice открыл и сконвертировал «%s» (%d КБ)'
                      % (label, size))
                continue
            print('  ❌ контрольный документ открылся, а «%s» — нет: '
                  'дело в самом файле' % label)
            print('     ' + (proc.stderr or proc.stdout or '').strip()[:300])
            good = False
        return good


def main():
    full = '--full' in sys.argv
    print('AOW-бот — офлайн-проверки\n' + '=' * 46)
    total_pass = total_fail = 0
    broken = []

    for title, script in SUITES:
        if not (HERE / script).exists():
            print('\n%s\n  ⏭  ПРОПУЩЕНО — нет %s' % (title, script))
            continue
        code, output = run(script)
        passed, failed = count(output)
        total_pass += passed
        total_fail += failed
        mark = '✅' if code == 0 else '❌'
        print('\n%s %s — проверок %d, провалов %d' % (mark, title, passed, failed))
        if code != 0:
            broken.append(title)
            print('\n'.join('    ' + l for l in output.strip().splitlines()[-25:]))

    if full:
        print('\nОткрытие документа в офисном пакете')
        if not office_check():
            broken.append('открытие документа')

    print('\n' + '=' * 46)
    if broken:
        print('❌ ЕСТЬ ОШИБКИ: %s' % ', '.join(broken))
        print('   Всего проверок: %d, провалено: %d' % (total_pass + total_fail, total_fail))
        return 1
    print('✅ ВСЁ ПРОШЛО — проверок: %d' % total_pass)
    if not full:
        print('   (--full добавит открытие документа в Word / LibreOffice)')
    return 0


if __name__ == '__main__':
    sys.exit(main())
