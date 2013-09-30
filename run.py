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
    load xls/xlsx file containing card-payment records
    """
    data = {}
    try:
        #
        sys.stderr.write("""loading "%s" ... """ % fn)
        book = xlrd.open_workbook(fn)
        sh = book.sheet_by_index(0)
        assert(sh.nrows > 1 and sh.ncols > 1)
        sys.stderr.write("done\n")
        #
        sys.stderr.write("""processing %d records ... """ % (sh.nrows - 1))
        cnt = 0
        for r in range(1, sh.nrows):  # skip header line
            tup = sh.row(r)
            # record schema: (card # text, payment float)
            card, value = str(tup[0].value).strip(), float(tup[1].value)
            # each card may have more than one payment records
            if not card in data:
                data[card] = []
            data[card].append(value)
            cnt += 1
            pass
        #
        assert(sh.nrows == cnt + 1)
        sys.stderr.write("done\n")
        pass
    except:
        sys.stderr.write("failed\n")
        raise
    return data


def payment2balance(seq):
    return 100.00 - sum(seq)


def save(data, fn, row_limit=65536, agg=payment2balance):
    """
    write aggregated per-card balance to xls file
    """
    try:
        #
        sys.stderr.write("""creating "%s" ... """ % fn)
        book = xlwt.Workbook(encoding='utf-8', style_compression=True)
        sys.stderr.write("done\n")
        #
        sys.stderr.write("""computing %d aggregates (with %s) """ %
                         (len(data), agg.__name__))
        style0 = xlwt.easyxf(num_format_str='@')
        style1 = xlwt.easyxf(num_format_str='0.00')
        sh = None
        cnt = 0
        for card in sorted(data):
            page = 1 + cnt / (row_limit - 1)
            row = 1 + cnt % (row_limit - 1)
            # start a new work sheet
            if row == 1:
                sh = book.add_sheet('Sheet%d' % page)
                sh.write(0, 0, '卡号')
                sh.write(0, 1, '余额')
                sh.col(0).width = 6500
                sh.col(1).width = 2000
                if page > 1:
                    sys.stderr.write("]")
                sys.stderr.write("[%d" % page)
            # write a record
            sh.write(row, 0, card, style0)
            sh.write(row, 1, agg(data[card]), style1)
            cnt += 1
            pass
        # make sure all records have been written
        assert(cnt == len(data))
        sys.stderr.write("] ... ")
        book.save(fn)
        sys.stderr.write("done\n")
        pass
    except:
        sys.stderr.write("failed\n")
        raise
    pass


def main(args):
    if len(args) <= 0:
        sys.stderr.write("usage: python %s <0.xls(x)> "
                         "[<1.xls(x)> [...]]\n" % __file__)
        return os.EX_USAGE
    for fi in args:
        fo = '.'.join(fi.split('.')[:-1] + ['agg', 'xls', ])
        save(load(fi), fo)
        pass
    return os.EX_OK

if __name__ == '__main__':
    sys.exit(main(sys.argv[1:]))
