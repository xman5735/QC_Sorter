import os

unc_path = r'\\lcc-fsqb-01.lcc.local\Shares\Alex\file.txt'

if os.path.exists(unc_path):
    with open(unc_path) as f:
        contents = f.read()
        print(contents)
else:
    print(f'Error: could not find file {unc_path}')
