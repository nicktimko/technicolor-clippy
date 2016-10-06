from __future__ import absolute_import, print_function

import re
import win32clipboard


def has_html():
    """
    Return True if there is a Html fragment in the clipboard..
    """
    cb = HtmlClipboard()
    return cb.has_html_format()


def get_html():
    """
    Return the Html fragment from the clipboard or None if there is no Html in the clipboard.
    """
    cb = HtmlClipboard()
    if cb.has_html_format():
        return cb.get_fragment()
    else:
        return None


def put_html(fragment):
    """
    Put the given fragment into the clipboard.
    Convenience function to do the most common operation
    """
    cb = HtmlClipboard()
    cb.put_fragment(fragment)


class HtmlClipboard(object):

    CF_HTML = None

    MARKER_BLOCK_OUTPUT = (
        "Version:1.0\r\n"
        "StartHTML:%09d\r\n"
        "EndHTML:%09d\r\n"
        "StartFragment:%09d\r\n"
        "EndFragment:%09d\r\n"
        "StartSelection:%09d\r\n"
        "EndSelection:%09d\r\n"
        "SourceURL:%s\r\n"
    )

    MARKER_BLOCK_EX = (
        "Version:(\S+)\s+"
        "StartHTML:(\d+)\s+"
        "EndHTML:(\d+)\s+"
        "StartFragment:(\d+)\s+"
        "EndFragment:(\d+)\s+"
        "StartSelection:(\d+)\s+"
        "EndSelection:(\d+)\s+"
        "SourceURL:(\S+)"
    )
    MARKER_BLOCK_EX_RE = re.compile(MARKER_BLOCK_EX)

    MARKER_BLOCK = (
        "Version:(\S+)\s+"
        "StartHTML:(\d+)\s+"
        "EndHTML:(\d+)\s+"
        "StartFragment:(\d+)\s+"
        "EndFragment:(\d+)\s+"
        "SourceURL:(\S+)"
    )
    MARKER_BLOCK_RE = re.compile(MARKER_BLOCK)

    DEFAULT_HTML_BODY = (
        "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">"
        "<HTML><HEAD></HEAD><BODY><!--StartFragment-->%s<!--EndFragment--></BODY></HTML>"
    )

    def __init__(self):
        self.html = None
        self.fragment = None
        self.selection = None
        self.source = None
        self.html_clipboard_version = None

    def get_cf_html(self):
        """
        Return the FORMATID of the HTML format
        """
        if self.CF_HTML is None:
            self.CF_HTML = win32clipboard.RegisterClipboardFormat("HTML Format")

        return self.CF_HTML

    def get_available_formats(self):
        """
        Return a possibly empty list of formats available on the clipboard
        """
        formats = []
        try:
            win32clipboard.OpenClipboard(0)
            cf = win32clipboard.EnumClipboardFormats(0)
            while (cf != 0):
                formats.append(cf)
                cf = win32clipboard.EnumClipboardFormats(cf)
        finally:
            win32clipboard.CloseClipboard()

        return formats

    def has_html_format(self):
        """
        Return a boolean indicating if the clipboard has data in HTML format
        """
        return (self.get_cf_html() in self.get_available_formats())


    def get_from_clipboard(self):
        """
        Read and decode the HTML from the clipboard
        """
        try:
            win32clipboard.OpenClipboard(0)
            src = win32clipboard.GetClipboardData(self.get_cf_html())
            #print src
            src = src.decode('utf-16')
            self.decode_clipboard_source(src)
        finally:
            win32clipboard.CloseClipboard()

    def decode_clipboard_source(self, src):
        """
        Decode the given string to figure out the details of the HTML that's on the string
        """
                    # Try the extended format first (which has an explicit selection)
        matches = self.MARKER_BLOCK_EX_RE.match(src)
        if matches:
            self.prefix = matches.group(0)
            self.html_clipboard_version = matches.group(1)
            self.html = src[int(matches.group(2)):int(matches.group(3))]
            self.fragment = src[int(matches.group(4)):int(matches.group(5))]
            self.selection = src[int(matches.group(6)):int(matches.group(7))]
            self.source = matches.group(8)
        else:
                    # Failing that, try the version without a selection
            matches = self.MARKER_BLOCK_RE.match(src)
            if matches:
                self.prefix = matches.group(0)
                self.html_clipboard_version = matches.group(1)
                self.html = src[int(matches.group(2)):int(matches.group(3))]
                self.fragment = src[int(matches.group(4)):int(matches.group(5))]
                self.source = matches.group(6)
                self.selection = self.fragment

    def get_html(self, refresh=False):
        """
        Return the entire Html document
        """
        if not self.html or refresh:
            self.get_from_clipboard()
        return self.html

    def get_fragment(self, refresh=False):
        """
        Return the Html fragment. A fragment is well-formated HTML enclosing the selected text
        """
        if not self.fragment or refresh:
            self.get_from_clipboard()
        return self.fragment

    def get_selection(self, refresh=False):
        """
        Return the part of the HTML that was selected. It might not be well-formed.
        """
        if not self.selection or refresh:
            self.get_from_clipboard()
        return self.selection

    def get_source(self, refresh=False):
        """
        Return the URL of the source of this HTML
        """
        if not self.selection or refresh:
            self.get_from_clipboard()
        return self.source

    def put_fragment(self, fragment, selection=None, html=None, source=None):
        """
        Put the given well-formed fragment of Html into the clipboard.

        selection, if given, must be a literal string within fragment.
        html, if given, must be a well-formed Html document that textually
        contains fragment and its required markers.
        """
        if selection is None:
            selection = fragment
        if html is None:
            html = self.DEFAULT_HTML_BODY % fragment
        if source is None:
            source = "file://HtmlClipboard.py"

        fragment_start = html.index(fragment)
        fragment_end = fragment_start + len(fragment)
        selection_start = html.index(selection)
        selection_end = selection_start + len(selection)
        self.put_to_clipboard(html, fragment_start, fragment_end, selection_start, selection_end, source)

    def put_to_clipboard(self, html, fragment_start, fragment_end, selection_start, selection_end, source="None"):
        """
        Replace the Clipboard contents with the given html information.
        """
        try:
            win32clipboard.OpenClipboard(0)
            win32clipboard.EmptyClipboard()
            src = self.encode_clipboard_source(html, fragment_start, fragment_end, selection_start, selection_end, source)
            #print src
            win32clipboard.SetClipboardData(self.get_cf_html(), src)
        finally:
            win32clipboard.CloseClipboard()

    def encode_clipboard_source(self, html, fragment_start, fragment_end, selection_start, selection_end, source):
        """
        Join all our bits of information into a string formatted as per the HTML format specs.
        """
                    # How long is the prefix going to be?
        dummy_prefix = self.MARKER_BLOCK_OUTPUT % (0, 0, 0, 0, 0, 0, source)
        len_prefix = len(dummy_prefix)

        prefix = self.MARKER_BLOCK_OUTPUT % (
            len_prefix,
            len(html) + len_prefix,
            fragment_start + len_prefix,
            fragment_end + len_prefix,
            selection_start + len_prefix,
            selection_end + len_prefix,
            source
        )
        return (prefix + html)


def dump_html():
    cb = HtmlClipboard()
    print("GetAvailableFormats()=%s" % str(cb.get_available_formats()))
    print("HasHtmlFormat()=%s" % str(cb.has_html_format()))
    if cb.has_html_format():
        cb.get_from_clipboard()
        print("prefix=>>>%s<<<END" % cb.prefix)
        print("htmlClipboardVersion=>>>%s<<<END" % cb.html_clipboard_version)
        print("GetSelection()=>>>%s<<<END" % cb.get_selection())
        print("GetFragment()=>>>%s<<<END" % cb.get_fragment())
        print("GetHtml()=>>>%s<<<END" % cb.get_html())
        print("GetSource()=>>>%s<<<END" % cb.get_source())


if __name__ == '__main__':
    def test_SimpleGetPutHtml():
        data = "<p>Writing to the clipboard is <strong>easy</strong> with this code.</p>"
        put_html(data)
        if get_html() == data:
            print("passed")
        else:
            print("failed")

    test_SimpleGetPutHtml()
    dump_html()
