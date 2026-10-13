# Third-Party Notices

OnyxPDF incorporates the following third-party components:

| Component | Version | License | Source |
|---|---|---|---|
| pdf.js (pdfjs-dist) | 4.10.38 | Apache-2.0 | https://mozilla.github.io/pdf.js/ |
| pdf-lib | 1.17.1 | MIT | https://pdf-lib.js.org |
| @pdf-lib/fontkit | 1.1.1 | MIT | https://github.com/Hopding/fontkit |
| pptxgenjs | 4.0.1 | MIT | https://gitbrent.github.io/PptxGenJS/ |
| fflate | 0.8.2 | MIT | https://github.com/101arrowz/fflate |
| ms | 2.1.3 | MIT | https://github.com/vercel/ms |
| http-proxy-agent | 9.1.0 | MIT | https://github.com/TooTallNate/proxy-agents |
| Electron | 33.4.11 | MIT | https://github.com/electron/electron |
| electron-builder (build time) | 25.1.8 | MIT | https://github.com/electron-userland/electron-builder |
| pdf.js standard fonts | same as pdf.js resources | BSD 3-Clause / SIL OFL 1.1 | https://mozilla.github.io/pdf.js/ |

## Notes

- Bundled copies of pdf.js, pdf-lib, @pdf-lib/fontkit, and pptxgenjs ship
  under `src/renderer/vendor/`. License files for the corresponding npm
  packages are included under `node_modules/` and in packaged builds that
  bundle those package directories.
- CMap resources under `src/renderer/vendor/cmaps/` are provided by the
  pdf.js project (Copyright Adobe). Their redistribution license is included
  at `src/renderer/vendor/cmaps/LICENSE`.
- Standard font resources under `src/renderer/vendor/standard_fonts/` are
  bundled from pdf.js. Their licenses are included in the same directory:
  `LICENSE_FOXIT` (PDFium authors, BSD 3-Clause) and `LICENSE_LIBERATION`
  (Liberation Sans, SIL Open Font License 1.1).
- The Apache-2.0 license text ships with `node_modules/pdfjs-dist/LICENSE`
  and is included in packaged builds that bundle the pdfjs-dist package.
- Where a bundled vendor file does not carry the full license text of its
  MIT-licensed upstream project, the following notice applies, together with
  the respective upstream copyright holders:

> Permission is hereby granted, free of charge, to any person obtaining a
> copy of this software and associated documentation files (the "Software"),
> to deal in the Software without restriction, including without limitation
> the rights to use, copy, modify, merge, publish, distribute, sublicense,
> and/or sell copies of the Software, and to permit persons to whom the
> Software is furnished to do so, subject to the following conditions:
>
> The above copyright notice and this permission notice shall be included in
> all copies or substantial portions of the Software.
>
> THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
> IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
> FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
> AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
> LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
> FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
> DEALINGS IN THE SOFTWARE.
