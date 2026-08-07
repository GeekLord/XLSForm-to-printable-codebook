"""Throwaway harness: exercise the unchanged generate_pdf() using locally installed Chrome.

Playwright's own Chromium build cannot be downloaded on this machine (TLS interception),
so launch() is patched only to supply an executable path. The PDF code path is untouched.
"""
import asyncio
import importlib.util
import os
import sys

CHROME = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

import playwright.async_api._generated as generated

_original_launch = generated.BrowserType.launch


async def _launch_with_local_chrome(self, **kwargs):
    kwargs.setdefault('executable_path', CHROME)
    return await _original_launch(self, **kwargs)


generated.BrowserType.launch = _launch_with_local_chrome

spec = importlib.util.spec_from_file_location('xtp', '../xlsform_to_printable.py')
xtp = importlib.util.module_from_spec(spec)
spec.loader.exec_module(xtp)

html_path, pdf_path = sys.argv[1], sys.argv[2]
asyncio.run(xtp.generate_pdf(html_path, pdf_path))
size = os.path.getsize(pdf_path)
with open(pdf_path, 'rb') as handle:
    head = handle.read(8)
print(f'{pdf_path}: {size} bytes, header {head!r}')
