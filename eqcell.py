import argparse
from xl import ExcelSheet

def cli():
    """Command line interface to XlSheet(filepath, sheet, anchor).save()"""
    
    parser = argparse.ArgumentParser(description='Command line interface to XlSheet(filename, sheet, anchor).save()',
                                     formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('filename',                        help='filename or path to .xls file')
    parser.add_argument('sheet',  nargs='?', default=1,  help='sheet name or sheet index starting at 1')
    parser.add_argument('anchor', nargs='?', default='A1', help='reference to upper-left corner of data block, defaults to A1')
    
    args = parser.parse_args()

    filename = args.filename
    anchor = args.anchor

    try:
        sheet = int(args.sheet)
    except ValueError:
        sheet = args.sheet
   
    xl = ExcelSheet(filename, sheet, anchor).save()
    xl.echo()

if __name__ == "__main__":
    cli()