import sys
import time

from pygments import highlight
from pygments.lexers import PythonLexer
from pygments.formatters import HtmlFormatter
import win32clipboard

from .html_clipboard import HtmlClipboard, put_html


def highlight_python(snippet):
    return highlight(snippet, PythonLexer(), HtmlFormatter(noclasses=True))


def main():
    win32clipboard.OpenClipboard()
    snippet = win32clipboard.GetClipboardData()

    highlighted = highlight_python(snippet)

    print(highlighted)
    # put_html(highlighted)

    time.sleep(5)


if __name__ == '__main__':
    sys.exit(main())
