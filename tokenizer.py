# *coding = utf-8

from typing import NamedTuple
import re

class Token(NamedTuple):
    type: str
    value: str
    line: int
    column: int

def tokenize(code):
        keywords = {'非常态化任务','监控、巡检','安全整改','变更发版','故障处理'}
        token_specification = [
                ('BINGO',r'[\u4e00-\u9fa5]'),
                ('END',r'END'),
                ('SKIP',r'None'),
                ]
        tok_regex = '|'.join('(?P<%s>%s)'% pair for pair in token_specification)
        line_num = 1
        line_start = 0

        for mo in re.finditer(tok_regex,code):
            kind = mo.lastgroup
            value = mo.group()
            column = mo.start() - line_start

            if kind == 'NUMBER':
                value = float(value) if '.' in value else int (value)
            elif kind == 'ID' and value in keywords:
                kind = value
            elif kind == 'NEWLINE':
                line_start = mo.end()
                line_num += 1
                continue
            elif kind == 'SKIP':
                continue
            elif kind == 'MISMATCH':
                raise RuntimeError(f'{value!r} unexpected on line {line_num}')

            yield Token(kind,value,line_num,column)

if __name__ == '__main__':

    statement = '''
    IF quantity THEN
        total := total + price * quantity;
        tax := price * 0.05
    ENDIF
    '''

    for token in tokenize(statement):
        print(token)
