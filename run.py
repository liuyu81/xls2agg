#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Copyright (C) 2013, LIU Yu <liuyu@opencps.net>
# All rights reserved.
#

import os
import sys
import xlrd
import xlwt


def load(fn):
    """
    load xls/xlsx file containing payment records
    """
    data = {}
    try:
        sys.stderr.write("""loading "%s" ... """ % fn)
        book = xlrd.open_workbook(fn)
        sh = book.sheet_by_index(0)
        assert(sh.nrows > 1 and sh.ncols > 1)
        sys.stderr.write("done\n")
        #
        sys.stderr.write("""reading %d records ... """ % (sh.nrows - 1))
        cnt = 0
        for r in range(1, sh.nrows):  # skip header line
            tup = sh.row(r)
            card, value = str(tup[0].value).strip(), float(tup[1].value)
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
    """
    write per-card balance to xls file
    """
    try:
        sys.stderr.write("""creating "%s" ... """ % fn)
        book = xlwt.Workbook(encoding='utf-8', style_compression=True)
        sys.stderr.write("done\n")
        #
        sys.stderr.write("""writing %d records ... """ % len(data))
        style0 = xlwt.easyxf(num_format_str='@')
        style1 = xlwt.easyxf(num_format_str='0.00')
        limit = 65535
        sh = None
        cnt = 0
        for card in sorted(data):
            page = cnt / limit
            row = cnt % limit
            if row == 0:
                sh = book.add_sheet('Sheet%d' % (page + 1))
                sh.write(0, 0, '卡号')
                sh.write(0, 1, '余额')
                sh.col(0).width = 6500
                sh.col(1).width = 2000
                sys.stderr.write("[%d]" % (page + 1))
            sh.write(row + 1, 0, card, style0)
            sh.write(row + 1, 1, 100.00 - sum(data[card]), style1)
            cnt += 1
        assert(cnt == len(data))
        book.save(fn)
        sys.stderr.write("\n")
        pass
    except:
        sys.stderr.write("failed\n")
        raise
    pass


def main(args):
    if len(args) <= 0:
        sys.stderr.write("usage: python %s <0.xls|xlsx> "
                         "[<1.xls|xlsx> [...]]\n" % __file__)
        return os.EX_USAGE
    for fi in args:
        fo = '.'.join(fi.split('.')[:-1] + ['agg', 'xls', ])
        save(load(fi), fo)
        pass
    return os.EX_OK

if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
