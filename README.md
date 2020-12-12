# PackageParser
Process archives containing forensic artifacts. Script can target individual archive or a directory containing multiple archives. PackageParser will extract the package and locate artifacts contained in the package for parsing (doesn't rely on known file paths). Output is written to a folder specified at the command-line.

If the --search option is selected, output files will be searched for patterns contained in regex.txt (located in search folder). PackageParser will accept regex or simple strings to search for in output files and will write a new CSV with matches to the output folder. 

## Example usage
`python PackageParser.py -s \path\to\source_dir -o \path\to\out_dir -p <password> --search`
