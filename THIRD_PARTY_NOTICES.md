# third-party notices

이 저장소는 아래 제3자 구성요소를 포함합니다.

## security_module/FilePathCheckerModule.dll

한글(HWP) COM 자동화 시 파일 접근마다 표시되는 보안 확인 대화상자를 억제하는 보안 모듈입니다.
한국어 HWP 자동화 생태계에서 널리 재배포되어 온 구성요소이며, 본 저장소는
[pyhwpx](https://github.com/martiniifun/pyhwpx) (MIT License) 배포본을 그대로 포함합니다.

```
MIT License

Permission is hereby granted, free of charge, to any person obtaining
a copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software, and to
permit persons to whom the Software is furnished to do so, subject to
the following conditions:

The above copyright notice and this permission notice shall be
included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT.
IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY
CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT,
TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE
SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
```

바이너리 포함을 원하지 않는 경우 `security_module/` 폴더를 삭제해도 서버는 정상 동작합니다.
다만 한글이 파일에 접근할 때마다 확인 대화상자가 떠서 자동화가 중단되므로,
직접 보안 모듈을 준비해 아래 레지스트리 값에 절대 경로를 등록해야 합니다.

```
HKCU\Software\HNC\HwpAutomation\Modules\FilePathCheckerModule
```

## 원본 저장소

본 저장소는 [crowwan/hwp-mcp-advanced-custom](https://github.com/crowwan/hwp-mcp-advanced-custom)
에서 갈라져 나온 포크입니다. 원본 저장소에는 명시된 라이선스가 없으며,
원본 코드의 권리는 원저작자에게 있습니다.
