#!/usr/bin/env python
# History:
#   2020-04-07 - initial development
from pathlib import Path
import re

__app_name__ = 'xlm_parse'
__app_home__ = Path(__file__).parent
__author__ = 'Harli Aquino <maharlito.aquino@cyren.com>'
__github__ = 'https://github.com/d34ddr34m3r'
__copyright__ = '(c) 2020 Cyren, Inc.'
__version__ = 0.1

cell_id_ptn = re.compile(r'\[(\$[A-Z]+\$\d+) len=\d+\]', re.S)
cell_value_ptn = re.compile(r'\[\[ "?(.*?)"? \]\]', re.S)
builtin_ptn = re.compile(r'Builtin - (\w+)', re.S)
run_ptn = re.compile(r'=(?:RUN|GOTO)\((.*)\)', re.S)
last_caller = None
caller = None
call_stack = []


def xlm_parse(lines, show_formula):
    global caller, last_caller
    max_empty_cells = 10
    cells = {}
    result = []

    for line in lines:
        try:
            if 'LABEL :' in line and 'Builtin -' in line:
                cell_ref = cell_value_ptn.search(line)
                if cell_ref is not None:
                    caller = cell_ref.group(1).replace('=', '')
                    caller_type = builtin_ptn.search(line).group(1)
            elif 'FORMULA :' in line:
                cell_id = cell_id_ptn.search(line).group(1)
                cells[cell_id] = {'formula': cell_value_ptn.search(line).group(1)}
            elif 'STRING :' in line:
                cells[cell_id]['string'] = cell_value_ptn.search(line).group(1)
        except Exception as error:
            # __logger__.error(error)
            pass

    def next_cell():
        global caller, last_caller, call_stack
        # __logger__.warning('SKIPPING TO NEXT CELL')
        while True:
            col, row = caller[1:].split('$')
            last_caller = '${}${}'.format(col, int(row) + 1)
            if last_caller in call_stack:
                caller = last_caller
                continue
            break
        call_stack.append(last_caller)
        return last_caller

    if caller is not None:
        message = '[{}] ={}'.format(caller_type, caller)
        # __logger__.debug(message)
    else:
        message = 'No auto-executable cell found. Using first Formula as entry-point.'
        # __logger__.error(message)
        for cell_id, cell_prop in cells.items():
            if 'formula' in cell_prop:
                caller = cell_id
                break
    result.append('===============================================================================')
    result.append(' {} {} by Harli Aquino <github.com/d34ddr34m3r>'.format(__app_name__, __version__))
    result.append('-------------------------------------------------------------------------------')
    result.append(message)
    halt = False
    while caller is not None:
        if not max_empty_cells:
            message = 'Maximum number of empty cells reached. Parser halted.'
            result.append(message)
            # __logger__.warning()
            break
        if caller in cells:
            this_caller = caller
            if 'formula' not in cells[caller]:
                break
            formula_string = None
            formula = cells[caller]['formula']
            callee = run_ptn.search(formula)
            if 'string' in cells[caller]:
                formula_string = cells[caller]['string']
            if callee is not None:
                caller = callee.group(1)
                if '~' in caller:
                    caller = caller.replace('~', '$')
            elif '=CALL' in formula or '=FORMULA' in formula:
                raw_formula = formula
                for cell_ref in re.compile(r'([~$][A-Z]+[~$]\d+)', re.S).findall(formula):
                    if cell_ref.replace('~', '$') in cells and 'string' in cells[cell_ref.replace('~', '$')]:
                        formula = formula.replace("{} ".format(cell_ref), '"{}" '.format(cells[cell_ref.replace("~", "$")]["string"]))
                        formula = formula.replace("{},".format(cell_ref), '"{}",'.format(cells[cell_ref.replace("~", "$")]["string"]))
                if '=FORMULA' in formula:
                    formula = formula.replace('" & "', '')
                caller = next_cell()
            elif formula in ['=HALT()', '=RETURN()']:
                halt = True
            else:
                caller = next_cell()
            if show_formula:
                formula_text = formula.replace("~", "").replace("$", "").replace(" ", "")
                formula_text = re.sub(r'^=FORMULA\((.*?),[a-zA-Z]+\d+\)', r'=\1', formula_text)
                message = '{}{}'.format(formula_text, "[" + formula_string + "]" if formula_string is not None else "")
            else:
                message = '{}{}'.format(formula.replace("~", "").replace("$", "").replace(" ", ""), "[" + formula_string + "]" if formula_string is not None else "")
            # print(message)
            result.append(message)
            if halt:
                break
        else:
            caller = next_cell()
            max_empty_cells -= 1
    caller, last_caller, call_stack = None, None, []
    return result


def main():
    def getargs():
        from optparse import OptionParser
        parser = OptionParser(usage='%prog [options] <args>')
        parser.add_option('-a', '--auto_exec_cell', dest='auto_exec_cell', default=None,
                          help="auto-exec cell")
        parser.add_option('-s', '--show_formula', dest='show_formula', action='store_true', default=False,
                          help='show formula')
        parser.add_option('-v', '--version', dest='version', action='store_true', default=False,
                          help="show version information")
        return parser.parse_args()

    options, args = getargs()
    if options.version:
        print('{} v{:0.02f}'.format(__app_name__, __version__))
        return 0

    for arg in args:
        path = Path(arg).expanduser().absolute()
        lines = [line for line in path.read_text().splitlines(keepends=False) if line.startswith("'")]
        print('\n'.join(xlm_parse(lines, options.show_formula)))


if __name__ == '__main__':
    main()
# ~nuninuninu~
