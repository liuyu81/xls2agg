#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import sys
import xlrd
import xlwt


def load(fn):
    data = {}
    # load xls file
    try:
        sys.stderr.write("loading \"%s\" ... " % fn)
        book = xlrd.open_workbook(fn)
        sh = book.sheet_by_index(0)
        assert(sh.nrows > 1 and sh.ncols > 1)
        sys.stderr.write("done\n")
        #
        sys.stderr.write("reading %d records ... " % (sh.nrows -  1))
        cnt = 0
        for r in range(1, sh.nrows): # skip header line
            tup = sh.row(r)
            card, value = long(tup[0].value), float(tup[1].value)
            if not card in data:
                data[card] = []
            data[card].append(value)
            cnt += 1
        assert(sh.nrows == cnt + 1)
        sys.stderr.write("done\n")
        pass
    except:
        sys.stderr.write("failed\n")
        raise
    return data


def save(data, fn):
    # write xls file
    try:
        sys.stderr.write("creating \"%s\" ... " % fn)
        book = xlwt.Workbook(encoding='utf-8', style_compression=True)
        sh = book.add_sheet(u"""余额汇总""")
        sh.write(0, 0, u"""卡号""")
        sh.write(0, 1, u"""余额""")
        sh.col(0).width = 6500
        sh.col(1).width = 2000
        style0 = xlwt.easyxf(num_format_str='@')
        style1 = xlwt.easyxf(num_format_str='0.00')
        sys.stderr.write("done\n")
        #
        sys.stderr.write("writing %d records ... " % len(data))
        r = 0
        for card, value in data.iteritems():
            r += 1
            sh.write(r, 0, str(card), style0)
            sh.write(r, 1, 100.00 - sum(value), style1)
        assert(r == len(data))
        book.save(fn)
        sys.stderr.write("done\n")
        pass
    except:
        sys.stderr.write("failed\n")
        raise
    pass


def main(args):
    if len(args) > 0:
        for fi in args:
            fo = '.'.join(fi.split('.')[:-1] + ['agg', 'xls', ])
            save(load(fi), fo)
    else:
        sys.stderr.write("usage: python %s <0.xls|xlsx> [<1.xls|xlsx> [...]\n" % __file__)
    return 0

if __name__ == '__main__':
	sys.exit(main(sys.argv[1:]))
