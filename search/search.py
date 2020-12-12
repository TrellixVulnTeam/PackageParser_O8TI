import re
from pathlib import Path
import argparse
import csv
import sys
from datetime import datetime
from colorama import init, Fore
from alive_progress import alive_bar

init(autoreset=True)

examples = '''
python search.py -s \\path\\to\\directory\\withCSVs -o \\path\\to\\out --search (uses default)
python search.py -s \\path\\to\\directory\\withCSVs -o \\path\\to\\out --search yourfile.txt

'''
rgx_dict = {}
str_dict = {}


def find_hits(file_list, dict1, dict2):
    """
    Search parsed artifact output CSVs for regex/strings
    :param file_list: list of CSV files to search
    :param dict1: dictionary with regex
    :param dict2: dictionary with strings
    :return: matches, re compile errors
    """
    re_errors = []
    matches = []
    total_patterns = len(dict1) + len(dict2)

    print(Fore.LIGHTWHITE_EX + f'Found {len(file_list)} CSV files. Searching using {total_patterns} patterns...\n')
    with alive_bar(len(file_list), bar='smooth') as bar:
        for file in file_list:
            print(Fore.LIGHTWHITE_EX + f'{file.name}' + Fore.LIGHTGREEN_EX)
            for key, val in dict1.items():
                if key == '[a-zA-Z0-9/+=]{500}' and 'history' in file.name.lower():
                    pass
                else:
                    try:
                        rgx = re.compile(key)
                    except Exception as e:
                        re_errors.append(f'{key}: {e}')
                    else:
                        with file.open('r', encoding='utf-8', errors='replace') as fh:
                            for line in fh:
                                if rgx.search(line):
                                    row = [file.name, key, val, line]
                                    matches.append(row)

            for key, val in dict2.items():
                with file.open('r', encoding='utf-8', errors='replace') as fh:
                    for line in fh:
                        if key.lower() in line.lower():
                            row = [file.name, key, val, line]
                            matches.append(row)
            bar()

    return matches, re_errors


def write_csv(hit_list, out_path):
    """
    Writes CSV containing pattern matches
    :param hit_list: matched patterns from regex/string search
    :param out_path: path to write CSV
    :return: nothing
    """
    dt = datetime.now()
    outdir = Path(out_path).joinpath('SearchResults')

    if not outdir.exists():
        outdir.mkdir(parents=True, exist_ok=True)

    out_file = outdir / f'SearchResults_{dt.strftime("%Y%m%d%H%M%S")}.csv'
    header = ['Source File', 'Match', 'Description', 'Found in Line']

    with out_file.open('w', newline='', errors='replace') as fh:
        csv_writer = csv.writer(fh)
        csv_writer.writerow(header)

        for hit in hit_list:
            csv_writer.writerow(hit)

    print(
        Fore.LIGHTRED_EX + f'\nFound {len(hit_list)} hits ' + Fore.LIGHTWHITE_EX + f'Check {out_file.name} for details.')


def main():
    search_path = Path(args.source)
    # get all CSV files in path, sort on file name.
    files = [i for i in sorted(search_path.rglob('*.csv'), key=lambda j: j.name) if 'SearchResults' not in i.name]

    if len(files) > 0:
        print(Fore.LIGHTWHITE_EX + f'\nBeginning search. Using {args.search} as input file.')
        hit_list, rgx_errors = find_hits(files, rgx_dict, str_dict)
        rgx_errors = list(set(rgx_errors))

        if len(rgx_errors) > 0:
            for i in rgx_errors:
                print(Fore.LIGHTRED_EX + f'[x] ERROR compiling regex: {i}')

        if len(hit_list) > 0:
            write_csv(hit_list, args.out)

        else:
            print(Fore.YELLOW + '\nNo matches found.')
    else:
        print(Fore.LIGHTRED_EX + f'No CSV files found in {search_path}')


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='IOC finder', epilog=examples,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-s', '--source', type=str, required=True,
                        help='Full path to directory containing csv/text files')
    parser.add_argument('-o', '--out', type=str, required=True, help='Output directory for CSV with IOC hits')
    parser.add_argument('--search', type=str, action='store', nargs='?', required=True, const='regex.txt',
                        help=' file to use. must be placed in cwd. Default is regex.txt')
    args = parser.parse_args()
    if args.search:
        rgx_file = Path.cwd() / args.search
        if not rgx_file.is_file():
            sys.exit(f'\nCan\'t find {args.search}. This file should be in the same dir as autoRon.')
        else:
            with rgx_file.open() as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                try:
                    for row in reader:
                        if any(row):
                            if row[0] == '1':  # regex if int(row[0])
                                rgx_dict[row[1]] = row[2]
                            elif row[0] == '0':  # simple string
                                str_dict[row[1]] = row[2]
                except Exception as e:
                    # print(f'\n[x] Formatting issue with regex.txt {e}. Please inspect: {row}')
                    print(Fore.YELLOW + '\nFormatting issue with row in regex.txt: ' + Fore.LIGHTWHITE_EX + f'{e}')
                    print(Fore.YELLOW + 'Please inspect: ' + Fore.LIGHTWHITE_EX + f'{row}')
    main()
