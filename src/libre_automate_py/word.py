from __future__ import annotations

from ooodev.utils.file_io import FileIO
from ooodev.loader import Lo
from ooodev.write import WriteDoc
from ooodev.write import Write
from com.sun.star.util import XSearchable, XReplaceDescriptor, XReplaceable
from com.sun.star.text import XTextRange
from typing import Sequence
from officeLoader import OfficeLoader

class Word:
    def __init__(self, read_only: bool = False, filepath: str | None = None, visible: bool = True) -> None:
        self._read_only = read_only
        self._filepath = filepath
        self._visible = visible
        self.doc = None

        try:
            office_loader = OfficeLoader()
            loader = office_loader.get_loader()
            if self._filepath:
                self._input_fnm = FileIO.get_absolute_path(self._filepath)
                self.doc = WriteDoc.open_doc(fnm=self._input_fnm, loader=loader, visible=self._visible)
            else:
                self.doc = WriteDoc.create_doc(visible=True)
        except Exception:
            Lo.close_office()
            raise

    def save(self, save_path: str | None = None) -> None:
        if not self.doc:
            raise RuntimeError("No document to save.")

        path = save_path or self._filepath
        if not path:
            raise ValueError("No file path specified for saving.")
        out_file = FileIO.get_absolute_path(path)
        _ = FileIO.make_directory(out_file)
        out_fnm = out_file
        try:
            self.doc.save_doc(fnm=out_fnm)
        except Exception:
            Lo.close_office()
            raise

    def close(self) -> None:
        self.doc.close_doc()


    def get_content_text(self) -> str:
        # iterate through the document contents, printing all the text portions in each paragraph

        text_doc = Write.get_text_doc(doc=self.doc)
        cursor = Write.get_cursor(text_doc)
        text = Write.get_all_text(cursor)
        return text

    def italicize_all(self, phrase: str) -> int:
        # cursor = Write.get_view_cursor(doc) # can be used when visible
        cursor = self.doc.get_cursor()
        cursor.goto_start()
        page_cursor = self.doc.get_view_cursor()
        result = 0
        try:
            searchable = self.doc.qi(XSearchable, True)
            search_desc = searchable.createSearchDescriptor()
            print(f"Searching for all occurrences of '{phrase}'")
            phrase_len = len(phrase)
            search_desc.setSearchString(phrase)

            matches = searchable.findAll(search_desc)
            result = matches.getCount()

            print(f"No. of matches: {result}")

            for i in range(result):
                match_tr = Lo.qi(XTextRange, matches.getByIndex(i))
                if match_tr is not None:
                    cursor.goto_range(match_tr, False)
                    print(f"  - found: '{match_tr.getString()}'")
                    print(f"    - on page {page_cursor.get_page()}")
                    cursor.goto_start(True)
                    print(
                        f"    - starting at char position: {len(cursor.get_string()) - phrase_len}"
                    )

        except Exception:
            raise
        return result

    def replace_words(self, old_words: Sequence[str], new_words: Sequence[str]) -> int:
        replace_n = 0

        for i in range(len(old_words)):
            print(f'replace {old_words[i]} to {new_words[i]}')
            replace_n = replace_n + self.replace_word(old_words[i], new_words[i])
        return replace_n

    def replace_word(self, old_word: str, new_word: str) -> int:
        replaceable = self.doc.qi(XReplaceable, True)
        replace_desc = Lo.qi(XReplaceDescriptor, replaceable.createSearchDescriptor())
        replace_desc.setSearchString(old_word)
        replace_desc.setReplaceString(new_word)
        return replaceable.replaceAll(replace_desc)

