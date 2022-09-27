import sys
import argparse
import pandas
import ctypes
import subprocess
from pathlib import Path
from colorama import init, Fore
from datetime import datetime
from zipfile import ZipFile
import pyfiglet
import tarfile
import csv
from search.search import write_csv, find_hits

init(autoreset=True)

example_text = '''
Examples:

Extract and process all packages in a directory:
python PackageParser.py -s \\path\\to\\archives -o \\path\\to\\out -p <password>

Extract and process an individual package and search output:
python PackageParser.py -s \\path\\to\\archive.7z -o \\path\\to\\out -p <password> --search

'''


class PackageParser:
    toolPath = Path.cwd() / 'tools'

    def __init__(self, source, out_dir, password=None, search=None):
        self.source = source
        self.password = password
        self.search = search

        if self.source.suffix == '.7z':
            self.package = Path(str(self.source)[:str(self.source).index('.7z')])
        elif self.source.suffix == '.zip':
            self.package = Path(str(self.source)[:str(self.source).index('.zip')])
        else:
            self.package = Path(str(self.source)[:str(self.source).index('.tar')])

        self.out_dir = Path(out_dir) / self.package.name

        if not self.out_dir.exists():
            self.out_dir.mkdir(parents=True, exist_ok=True)

        if self.search:
            self.rgx_dict = {}
            self.str_dict = {}
            self.rgx_file = Path.cwd() / 'search' / self.search

        print(Fore.LIGHTGREEN_EX + f'\nCreating PackageParser object for package: '
                                   f'' + Fore.LIGHTWHITE_EX + f'{self.source.name}\n')

    def searcher(self):
        """search parsed output for regex/strings"""
        if not self.rgx_file.is_file():
            print(Fore.LIGHTRED_EX + f'\nCan\'t find {self.rgx_file.name}. '
                                     f'This file should be placed in the search folder')
        else:
            with self.rgx_file.open() as csvfile:
                reader = csv.reader(csvfile, delimiter=';')
                try:
                    for row in reader:
                        if any(row):  # we want non blank rows
                            if row[0] == '1':  # regex
                                self.rgx_dict[row[1]] = row[2]  # {regex: description}
                            elif row[0] == '0':  # string
                                self.str_dict[row[1]] = row[2]  # {string: description}
                except Exception as e:
                    print(Fore.YELLOW + f'\nFormatting issue with row in {self.rgx_file.name}: ' +
                          Fore.LIGHTWHITE_EX + f'{e}')
                    print(Fore.YELLOW + 'Please inspect: ' + Fore.LIGHTWHITE_EX + f'{row}')

            files = [i for i in sorted(self.out_dir.rglob('*.csv'), key=lambda j: j.name) if
                     'SearchResults' not in i.name]

            if len(files) > 0:
                print(Fore.LIGHTWHITE_EX + f'\nSearch options selected. Using {self.rgx_file.name} as input file.')
                hit_list, rgx_errors = find_hits(files, self.rgx_dict, self.str_dict)

                rgx_errors = list(set(rgx_errors))
                if len(rgx_errors) > 0:
                    for i in rgx_errors:
                        print(Fore.LIGHTRED_EX + f'[x] ERROR compiling regex: {i}')

                if len(hit_list) > 0:
                    write_csv(hit_list, self.out_dir)
                else:
                    print(Fore.YELLOW + '\nNo matches found.')
            else:
                print(Fore.YELLOW + f'\nNo CSV files found in {self.out_dir}')

    def logger(self, lev, msg):
        line = str(datetime.now().replace(microsecond=0)) + ' | ' + lev + ': ' + msg
        log_path = self.out_dir / 'PackageParser.log'

        if lev == 'SUCCESS':
            print(Fore.LIGHTGREEN_EX + line)
        elif lev == 'NOTICE':
            print(Fore.LIGHTYELLOW_EX + line)
        elif lev == 'ERROR':
            print(Fore.LIGHTRED_EX + line)
        elif lev == 'INFO':
            print(Fore.LIGHTMAGENTA_EX + line)
        elif lev == 'DONE':
            print(Fore.LIGHTCYAN_EX + line)
        try:
            with log_path.open('a', encoding='utf-8', newline='') as fh:
                fh.write(line + '\r\n')
        except Exception as e:
            print(Fore.LIGHTRED_EX + f'\nERROR: {e}')

    def run_simp_command(self, command):
        """run subprocess and redirect console output to log"""
        ez_log = self.out_dir / 'tools.log'
        with ez_log.open('a') as fh:
            spr = subprocess.run(command, stdout=fh, stderr=fh, timeout=300)
            spr.check_returncode()

    def run_command(self, command, bin_path, artifact, out_path):
        """run subprocess and redirect console output to log"""
        ez_log = self.out_dir / 'tools.log'
        with ez_log.open('a') as fh:
            try:
                self.logger('INFO', f'Found {artifact}. Running {bin_path.name}')
                spr = subprocess.run(command, stdout=fh, stderr=fh, timeout=1200)
                spr.check_returncode()
                self.logger('SUCCESS', f'{artifact} output written to {out_path}')
            except subprocess.CalledProcessError as e:
                self.logger('ERROR', str(e))
            except subprocess.TimeoutExpired:
                self.logger('ERROR', f'{bin_path.name} exceeded 20 minute timeout')

    def extract_sevenzip(self):
        """extract password protected 7zip archives"""
        seven_zip = PackageParser.toolPath / 'sevenZip/7za.exe'
        command = [str(seven_zip), 'x', '-spe', '-o' + str(self.source.parent), '-p' + self.password,
                   str(self.source), '-aoa']
        try:
            self.logger('INFO', f'Extracting 7zip: {self.source.name}')
            self.run_command(command)
            self.logger('SUCCESS', f'Extracted 7zip: {self.source.name}')
            self.source.unlink()
            self.logger('INFO', f'Deleted 7zip Archive: {self.source.name}')
        except subprocess.CalledProcessError:
            self.logger('ERROR', 'Problem extracting 7zip. Check tools.log for details.')
            print(Fore.LIGHTRED_EX + '\nIt was probably your password.')
            sys.exit()

    def extract_tar(self):
        """extract gzipped TAR package"""
        try:
            self.logger('INFO', f'Extracting tar file: {self.source.name}')
            with tarfile.open(str(self.source)) as tf:
                tf.extractall(str(self.package))
            self.logger('SUCCESS', f'Extracted tar file: {self.source.name}')
            self.source.unlink()
            self.logger('INFO', f'Deleted tar file: {self.source.name}')
        except Exception as e:
            self.logger('ERROR', 'Problem extracting tar file: ' + str(e))
            sys.exit()

    def extract_zipfile(self):
        """extract .zip no password"""
        try:
            self.logger('INFO', f'Extracting zip file: {self.source.name}')
            with ZipFile(str(self.source), 'r') as zf:
                zf.extractall(str(self.package))
            self.logger('SUCCESS', f'Extracted zip file: {self.source.name}')
            self.source.unlink()
            self.logger('INFO', f'Deleted zip file: {self.source.name}')
        except Exception as e:
            self.logger('ERROR', 'Problem extracting zip file: ' + str(e))
            sys.exit()

    def convert_csv(self):
        """convert JSON files to CSV"""
        query_results = self.package / 'QueryResults'
        out_dir = self.out_dir / 'QueryResults'
        empty = not bool(list(query_results.glob('*.json')))

        if not empty:
            self.logger('INFO', 'Converting Query Results to CSV')
            if not out_dir.exists():
                out_dir.mkdir(parents=True, exist_ok=True)
            for i in query_results.iterdir():
                out_file = out_dir.joinpath(str(i.stem) + '.csv')
                try:
                    with i.open(encoding='utf-8-sig', errors='replace') as fh:
                        df = pandas.read_json(fh)
                    df.to_csv(str(out_file), encoding='utf-8', index=False)
                    # self.logger('SUCCESS', f'Converted: {i}')
                except Exception as e:
                    self.logger('ERROR', f'Problem converting: {i} : {str(e)}')
            self.logger('SUCCESS', f'JSON 2 CSV Output written to: {out_dir}')
        else:
            self.logger('NOTICE', f'No QueryResults found in package: {self.package.name}. Skipping...')

    def mft_parse(self):
        """find and parse $MFT and UsnJrnl"""
        mftecmd = PackageParser.toolPath / 'MFTECmd.exe'
        mft_list = list(self.package.rglob('$MFT'))
        j_list = list(self.package.rglob('$J'))

        if mft_list:
            mft_out = self.out_dir / 'Filesystem/MFT'
            for mft in mft_list:
                command = [str(mftecmd), '-f', '"' + str(mft) + '"', '--csv', str(mft_out)]
                self.run_command(command, mftecmd, '$MFT', str(mft_out))

        if j_list:
            j_out = self.out_dir / 'Filesystem/UsnJrnl'
            for j in j_list:
                command = [str(mftecmd), '-f', '"' + str(j) + '"', '--csv', str(j_out)]
                self.run_command(command, mftecmd, '$J', str(j_out))

    def shim_parse(self):
        """find and parse SYSTEM hive"""
        ez_shim = PackageParser.toolPath / 'AppCompatCacheParser.exe'
        empty = not bool(list(i for i in self.package.rglob('SYSTEM') if i.is_file()))

        if not empty:
            shim_path = list(i for i in self.package.rglob('SYSTEM') if i.is_file())
            shim_out = self.out_dir / 'ProgramExecution/Shimcache'
            command = [str(ez_shim), '-f', '"' + str(shim_path[0]) + '"', '--csv', str(shim_out), '--nl']
            self.run_command(command, ez_shim, 'SYSTEM Hive (shimcache', shim_out)
        else:
            self.logger('NOTICE', f'No SYSTEM file found in package: {self.package.name}. Skipping...')

    def amcache_parse(self):
        """find and parse Amcache"""
        ez_amc = PackageParser.toolPath / 'AmcacheParser.exe'
        empty = not bool(list(self.package.rglob('Amcache.hve')))

        if not empty:
            amc_path = self.package.rglob('Amcache.hve')
            amc_out = self.out_dir / 'ProgramExecution/Amcache'
            command = [str(ez_amc), '-f', '"' + str(next(amc_path)) + '"', '--csv', str(amc_out), '--nl']
            self.run_command(command, ez_amc, 'Amcache.hve', amc_out)
        else:
            self.logger('NOTICE', f'No Amcache.hve found in package: {self.package.name}. Skipping...')

    def rfc_parse(self):
        """find and parse RecentFileCache"""
        ez_rfc = PackageParser.toolPath / 'RecentFileCacheParser.exe'
        empty = not bool(list(self.package.rglob('RecentFileCache.bcf')))

        if not empty:
            rfc_path = self.package.rglob('RecentFileCache.bcf')
            rfc_out = self.out_dir / 'ProgramExecution/RecentFileCache'
            command = [str(ez_rfc), '-f', '"' + str(next(rfc_path)) + '"', '--csv', str(rfc_out)]
            self.run_command(command, ez_rfc, 'RecentFileCache.bcf', rfc_out)
        else:
            self.logger('NOTICE', f'No RecentFileCache.bcf found in package: {self.package.name}. Skipping...')

    def prefetch_parse(self):
        """find and parse Prefetch files"""
        pecmd = PackageParser.toolPath / 'PECmd.exe'
        empty = not bool(list(self.package.rglob('*.pf')))

        if not empty:
            pf_dir = list(self.package.rglob('*.pf'))
            pf_out = self.out_dir / 'ProgramExecution/Prefetch'
            command = [str(pecmd), '-d', '"' + str(pf_dir[0].parent) + '"', '--csv', str(pf_out)]
            self.run_command(command, pecmd, 'Prefetch files', pf_out)
        else:
            self.logger('NOTICE', f'No Prefetch files found in package: {self.package.name}. Skipping...')

    def reg_parse(self):
        """find and parse registry hive files"""
        recmd = PackageParser.toolPath / 'RECmd/RECmd.exe'
        reg_files = ['SYSTEM', 'SECURITY', 'SOFTWARE', 'SAM', 'NTUSER.DAT', 'UsrClass.DAT']
        empty = not bool(list(i for i in self.package.rglob('*') if i.name in reg_files))

        if not empty:
            reg_out = self.out_dir / 'Registry'
            batch_mc = PackageParser.toolPath / 'RECmd/RECmd_Batch_MC.reb'
            batch_command = [str(recmd), '-d', '"' + str(self.package) + '"', '--bn', str(batch_mc), '--csv',
                             str(reg_out) + '\\RECmdBatch', '--nl']

            reg_exe = PackageParser.toolPath / 'RECmd/AllRegExecutablesFoundOrRun.reb'
            exe_command = [str(recmd), '-d', '"' + str(self.package) + '"', '--bn', str(reg_exe), '--csv',
                           str(reg_out) + '\\RegEXEsFoundOrRun', '--nl']

            user_activity = PackageParser.toolPath / 'RECmd/UserActivity.reb'
            user_command = [str(recmd), '-d', '"' + str(self.package) + '"', '--bn', str(user_activity), '--csv',
                            str(reg_out) + '\\UserActivity', '--nl']
            try:
                self.logger('INFO', f'Parsing Registry Hives in package: {self.package.name}')
                self.run_command(batch_command)
                self.logger('SUCCESS', f'Registry output written to: {reg_out}\\RECmdBatch')

                self.logger('INFO', f'Parsing Registry Hives for EXEs '
                                    f'found or run in package: {self.package.name}')
                self.run_command(exe_command)
                self.logger('SUCCESS', f'EXEs found or run output written to: {reg_out}\\RegEXEsFoundOrRun')

                self.logger('INFO', f'Parsing Registry Hives for user activity '
                                    f'in package: {self.package.name}')
                self.run_command(user_command)
                self.logger('SUCCESS', f'User activity output written to: {reg_out}\\UserActivity')
            except subprocess.CalledProcessError as e:
                self.logger('ERROR', str(e))
            except subprocess.TimeoutExpired:
                self.logger('ERROR', f'{recmd.name} exceeded 5 minute timeout.')
        else:
            self.logger('NOTICE', f'No Registry Hives found in package: {self.package.name}. Skipping...')

    def winevt_parse(self):
        """find and parse event logs"""
        evtxecmd = PackageParser.toolPath / 'EvtxECmd/EvtxECmd.exe'
        empty = not bool(list(self.package.rglob('*.evtx')))

        if not empty:
            winevt_path = list(self.package.rglob('*.evtx'))
            winevt_out = self.out_dir / 'EventLogs'
            command = [str(evtxecmd), '-d', '"' + str(winevt_path[0].parent) + '"', '--csv', str(winevt_out)]
            self.run_command(command, evtxecmd, 'Event Logs', winevt_out)

    def shellbags_parse(self):
        """find and parse User registry hive files"""
        sbecmd = PackageParser.toolPath / 'SBECmd.exe'
        empty = not bool(list(self.package.rglob('*.DAT')))

        if not empty:
            sb_out = self.out_dir / 'FileFolderAccess/ShellBags'
            command = [str(sbecmd), '-d', '"' + str(self.package) + '"', '--csv', str(sb_out), '--nl']
            self.run_command(command, sbecmd, 'User hives (shellbags', sb_out)
        else:
            self.logger('NOTICE', f'No User Hives found in package: {self.package.name}. Skipping...')

    def lnk_parse(self):
        """find and parse LNK files"""
        lecmd = PackageParser.toolPath / 'LECmd.exe'
        empty = not bool(list(self.package.rglob('*.lnk*')))

        if not empty:
            lnk_out = self.out_dir / 'FileFolderAccess/LNKfiles'
            command = [str(lecmd), '-d', '"' + str(self.package) + '"', '--csv', str(lnk_out), '--all']
            self.run_command(command, lecmd, 'LNK files', lnk_out)

    def jumplist_parse(self):
        """find and parse Jump Lists"""
        jlecmd = PackageParser.toolPath / 'JLECmd.exe'
        empty = not bool(list(self.package.rglob('*Destinations-ms')))

        if not empty:
            jl_out = self.out_dir / 'FileFolderAccess/JumpLists'
            command = [str(jlecmd), '-d', '"' + str(self.package) + '"', '--csv', str(jl_out)]
            self.run_command(command, jlecmd, 'Jump Lists', jl_out)

    def run_all(self):
        start_time = datetime.now().replace(microsecond=0)
        if self.source.suffix == '.7z' and self.password is not None:
            self.extract_sevenzip()
        elif self.source.suffix == '.zip':
            self.extract_zipfile()
        else:
            self.extract_tar()
        self.convert_csv()
        self.mft_parse()
        self.amcache_parse()
        self.rfc_parse()
        self.shim_parse()
        self.prefetch_parse()
        self.reg_parse()
        self.winevt_parse()
        self.shellbags_parse()
        self.lnk_parse()
        self.jumplist_parse()
        self.logger('DONE', f'Processed {self.package.name} in {datetime.now().replace(microsecond=0) - start_time}')
        print(Fore.LIGHTGREEN_EX + '\nOutput written to: ' + Fore.LIGHTWHITE_EX + f'{self.out_dir}')
        if self.search:
            self.searcher()


def main():
    out_dir = args.out
    user_source = Path(args.source)

    if not ctypes.windll.shell32.IsUserAnAdmin() == 1:
        sys.exit(Fore.LIGHTRED_EX + 'Please rerun from an Administrative command prompt. Exiting')

    if not PackageParser.toolPath.exists():
        sys.exit(Fore.LIGHTRED_EX + '\nCan\'t find the tools folder. The tools folder should be placed in '
                                    'the same directory as PackageParser. Exiting.')

    if user_source.is_dir():
        if str(user_source) == out_dir:
            sys.exit(Fore.LIGHTRED_EX + '-s (--source) and -o (--out) cannot be the same directory. Exiting.')

        ext_glob = ['*.7z', '*.zip', '*.gz']
        archives = [a for a in [user_source.glob(e) for e in ext_glob] for a in a]

        if len(archives) > 0:
            pa = '\n'.join(map(str, archives))
            print(Fore.LIGHTGREEN_EX + f'\nFound {len(archives)} Package(s) '
                                       'in: ' + Fore.LIGHTWHITE_EX + f'{user_source}')
            print(Fore.LIGHTCYAN_EX + '\n' + pa)
            for archive in archives:
                if archive.suffix == '.7z' and not args.password:
                    sys.exit(Fore.LIGHTRED_EX + f'\nNo password provided for .7z. Exiting')
                else:
                    package = PackageParser(archive, out_dir, args.password, args.search)
                    package.run_all()
        else:
            sys.exit(Fore.LIGHTRED_EX + f'\nPath: {user_source} contains no packages. Exiting.')

    elif user_source.is_file():
        exts = ['.7z', '.zip', '.gz']
        if str(user_source.parent) == out_dir:
            sys.exit(Fore.LIGHTRED_EX + '\nOutput directory cannot be the same as the source file. Exiting.')

        if user_source.suffix in exts:
            print(Fore.LIGHTGREEN_EX + '\nFound package: ' + Fore.LIGHTWHITE_EX + f'{user_source}')

            if user_source.suffix == '.7z' and not args.password:
                sys.exit(Fore.LIGHTRED_EX + '\nNo password provided for .7z. Exiting.')
            else:
                package = PackageParser(user_source, out_dir, args.password, args.search)
                package.run_all()
        else:
            sys.exit(Fore.LIGHTRED_EX + f'\nWrong file type based on extension: {user_source.name}')
    else:
        sys.exit(Fore.LIGHTRED_EX + f'Invalid path {user_source}')


if __name__ == '__main__':
    package_print = pyfiglet.figlet_format('PackageParser', font='cosmic')
    print(Fore.LIGHTGREEN_EX + '\n' + package_print)

    parser = argparse.ArgumentParser(description='Process archives containing LR artifacts', epilog=example_text,
                                     formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument('-p', '--password', type=str, help='archive password')
    parser.add_argument('--search', type=str, action='store', nargs='?', const='regex.txt',
                        help='input file to use. must be placed in search folder. Default is regex.txt')

    required_args = parser.add_argument_group('required arguments')
    required_args.add_argument('-s', '--source', type=str, required=True,
                               help='Full path to archive OR directory containing '
                                    'multiple archives.')
    required_args.add_argument('-o', '--out', type=str, required=True, help='Processed package output directory.')
    args = parser.parse_args()
    main()
