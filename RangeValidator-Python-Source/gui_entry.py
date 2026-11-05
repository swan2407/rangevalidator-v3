from rangevalidator.gui import main
if __name__=='__main__':
    import sys
    if len(sys.argv)>1 and sys.argv[1]=='--full-integration-test':
        import json, traceback
        from pathlib import Path
        from rangevalidator.development_probe import run
        try:
            run(sys.argv[2])
        except Exception as exc:
            Path(sys.argv[2]).write_text(json.dumps(dict(error=str(exc),traceback=traceback.format_exc(),cleanup_errors=getattr(exc,'cleanup_errors',[])),ensure_ascii=False,indent=2),encoding='utf-8')
            raise SystemExit(1)
    elif len(sys.argv)>1 and sys.argv[1]=='--smoke-test':
        import json
        from pathlib import Path
        import tkinter as tk
        import pythoncom,win32api,openpyxl
        from rangevalidator.gui import App
        root=tk.Tk();app=App(root)
        def finish():
            Path(sys.argv[2]).write_text(json.dumps(dict(gui=True,tabs=app.notebook.tabs().__len__(),python=sys.version,pythoncom=pythoncom.__file__,win32api=win32api.__file__,frozen=getattr(sys,'frozen',False)),ensure_ascii=False),encoding='utf-8')
            root.destroy()
        root.after(1000,finish)
        root.mainloop()
    elif len(sys.argv)>1 and sys.argv[1]=='--batch-read-test':
        from rangevalidator.batch_probe import run
        run(sys.argv[2])
    elif len(sys.argv)>1 and sys.argv[1]=='--diagnostic-test':
        from rangevalidator.diagnostic_probe import run
        run(sys.argv[2])
    elif len(sys.argv)>1 and sys.argv[1]=='--integration-test':
        import json
        from pathlib import Path
        from rangevalidator.selfcheck import run
        try:
            result=run(Path(sys.argv[2]).with_suffix(''))
        except Exception as exc:
            result=dict(error=str(exc))
        Path(sys.argv[2]).write_text(json.dumps(result,ensure_ascii=False),encoding='utf-8')
        raise SystemExit(int('error' in result))
    else:main()
