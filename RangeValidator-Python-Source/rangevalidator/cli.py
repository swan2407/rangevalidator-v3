import argparse
from dataclasses import asdict, is_dataclass
from datetime import date, datetime, time
import json
from pathlib import Path
import sys
from .analysis import analyze
from .design import discover, compare
from .readers import read_workbook
from .models import ReadFailure


def json_default(value):
    if is_dataclass(value):
        return asdict(value)
    if isinstance(value, (date, datetime, time)):
        return value.isoformat()
    if isinstance(value, (set, frozenset)):
        return sorted(value)
    raise TypeError(f'Unsupported JSON value: {type(value).__name__}')


def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, 'reconfigure'):
            stream.reconfigure(encoding='utf-8')
    parser = argparse.ArgumentParser(description='Offline read-only bit-field and Excel inspection. JSON to stdout; no workbook writes.')
    sub = parser.add_subparsers(dest='command', required=True)
    c = sub.add_parser('analyze')
    c.add_argument('source', type=Path)
    c.add_argument('--encoding', default='utf-8-sig')
    x = sub.add_parser('read')
    x.add_argument('xlsx')
    x.add_argument('--backend', choices=['openpyxl', 'com'], default='openpyxl')
    p = sub.add_parser('compare')
    p.add_argument('source', type=Path)
    p.add_argument('xlsx')
    p.add_argument('--encoding', default='utf-8-sig')
    p.add_argument('--backend', choices=['openpyxl', 'com'], default='openpyxl')
    sub.add_parser('gui')
    batch=sub.add_parser('batch')
    batch.add_argument('details',nargs='+')
    batch.add_argument('--header',action='append',required=True)
    batch.add_argument('--backend',choices=['auto','openpyxl','com'],default='auto')
    batch.add_argument('--encoding',default='utf-8-sig')
    args = parser.parse_args(argv)
    if args.command=='gui':
        from .gui import main as gui_main
        gui_main()
        return 0
    if args.command=='batch':
        from .workflow import run_comparison
        try:
            session=run_comparison(None,args.header,args.details,args.backend,args.encoding)
            print(json.dumps(dict(results=session.results,errors=session.errors,readers=session.readers),default=json_default,ensure_ascii=False,indent=2))
            return 0
        except Exception as exc:
            err=exc if isinstance(exc,ReadFailure) else ReadFailure('batch',exc)
            print(json.dumps({'error':err.as_dict()},ensure_ascii=False),file=sys.stderr)
            return 2
    stage = 'source.read'
    try:
        model = None
        if args.command in ('analyze', 'compare'):
            model = analyze(args.source.read_text(encoding=args.encoding), str(args.source.resolve()))
        if args.command == 'analyze':
            output = model
        else:
            stage = 'workbook.read'
            book = read_workbook(args.xlsx, args.backend)
            output = book if args.command == 'read' else dict(
                source_reviews=model.reviews,
                results=[compare(model, row) for row in discover(book)])
        stage = 'json.output'
        print(json.dumps(output, default=json_default, ensure_ascii=False, indent=2))
        return 0
    except Exception as exc:
        err = exc if isinstance(exc, ReadFailure) else ReadFailure(stage, exc)
        print(json.dumps({'error': err.as_dict()}, ensure_ascii=False), file=sys.stderr)
        return 2

