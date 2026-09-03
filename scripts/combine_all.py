import pathlib
import re
from pprint import pprint

def main():
    cd = pathlib.Path('.')


    combined_bas = pathlib.Path('all.bas')
    combined_bas.touch()
    combined_bas.write_text('')

    target_pattern = f'^(?!{combined_bas.name}$)\w+\.bas'
    script_files = sorted([f for f in cd.iterdir() if re.match(target_pattern, f.name)])


    head_text = 'Option Explicit\n\n' \
    'Dim shapePositions() As Variant\n' \
    'Dim ShapeDistanceX As Double\n' \
    'Dim ShapeDistanceY As Double\n' \
    '\n' \
    'Sub InitCustomTab()\n' \
    '    \' Ribbon onLoad can fire before any presentation exists (add-in / new PPT instance).\n' \
    '    \' Do not access ActivePresentation unconditionally or error 80048240 appears.\n' \
    '    Dim slideWidth As Double\n' \
    '    Dim slideHeight As Double\n' \
    '\n' \
    '    \' Fallback: standard widescreen (13.333" x 7.5" in points)\n' \
    '    slideWidth = 960#\n' \
    '    slideHeight = 540#\n' \
    '\n' \
    '    On Error Resume Next\n' \
    '    If Application.Presentations.Count > 0 Then\n' \
    '        slideWidth = ActivePresentation.PageSetup.SlideWidth\n' \
    '        slideHeight = ActivePresentation.PageSetup.SlideHeight\n' \
    '    End If\n' \
    '    On Error GoTo 0\n' \
    '\n' \
    '    ShapeDistanceX = slideWidth * 0.05\n' \
    '    ShapeDistanceY = slideHeight * 0.01\n' \
    'End Sub\n' \
    
    print(head_text)

    with open(combined_bas.name, 'a') as f:
        f.write(head_text)
        for file in script_files:
            print(file.name)
            f.write(file.read_text(encoding='utf-8'))
            f.write('\n')


if __name__ == "__main__":
    main()