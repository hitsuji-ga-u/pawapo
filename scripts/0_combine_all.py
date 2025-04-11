import pathlib
import re
from pprint import pprint

def main():
    cd = pathlib.Path('.')

    all_scripts = pathlib.Path('all.bas')
    all_scripts.touch()
    all_scripts.write_text('')

    target_pattern = f'^(?!({all_scripts.name}\.)|(init\.))\w+\.bas'
    target_files = sorted([f for f in cd.iterdir() if f.is_file() and f.suffix == '.bas'])

    # ヘッダーを読み込む　一番最初に読んでおく
    head_text = ''
    with open('init.bas', 'r', encoding='utf-8') as f:
        head_text = f.read()

    with open(all_scripts.name, 'a', encoding='utf-8') as f:
        f.write(head_text)
        for file in target_files:
            if re.match(target_pattern, file.name):
                print('〇 ', end='')
                f.write(file.read_text(encoding='utf-8'))
                f.write('\n')
            else:
                print('✗ ', end='')
            print(file.name)


if __name__ == "__main__":
    main()