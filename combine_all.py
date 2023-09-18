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


    head_text = 'Option Explicit\n\n' 

    with open(combined_bas.name, 'a') as f:
        f.write(head_text)
        for file in script_files:
            print(file.name)
            f.write(file.read_text(encoding='utf-8'))
            f.write('\n')


if __name__ == "__main__":
    main()