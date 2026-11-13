# -*- coding: utf-8 -*-
"""通过 WPS / Word COM 更新 docx 目录（仅 Windows）。

用法：
    python update_toc.py <docx 路径>
"""
import os
import sys
import time


def update_toc(path: str) -> bool:
    """依次尝试 WPS 与 Word，更新文档中的第一个目录域。"""
    try:
        import win32com.client
    except ImportError:
        print("未安装 pywin32，请在 Word/WPS 中右键目录 → 更新域")
        return False

    for progid in ("KWPS.Application", "Word.Application"):
        app = None
        try:
            app = win32com.client.DispatchEx(progid)
            try:
                app.Visible = False
            except Exception:
                pass
            try:
                app.DisplayAlerts = 0
            except Exception:
                pass
            doc = app.Documents.Open(os.path.abspath(path))
            if doc.TablesOfContents.Count > 0:
                doc.TablesOfContents(1).Update()
            doc.Save()
            doc.Close(True)
            print(f"目录已更新（{progid}）")
            return True
        except Exception as exc:
            print(f"{progid} 更新失败: {exc}")
        finally:
            if app is not None:
                try:
                    app.Quit()
                except Exception:
                    pass
            time.sleep(1)
    return False


if __name__ == "__main__":
    if len(sys.argv) < 2:
        sys.exit("用法: python update_toc.py <docx 路径>")
    ok = update_toc(sys.argv[1])
    if not ok:
        print("目录更新失败，请在 Word/WPS 中右键目录 → 更新域")
    sys.exit(0 if ok else 1)
