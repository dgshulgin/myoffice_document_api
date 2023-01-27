#!/usr/bin/env python3

import sys
import os
from MyOfficeSDKDocumentAPI import DocumentAPI as dapi

this = sys.modules[__name__]
this.app    = None
this.doc    = None



def main() -> int:
    num_args = len(sys.argv)
    if num_args < 2:
        return 1
    template = sys.argv[1]

    if this.app is None:
        this.app = dapi.Application()

    if this.doc is None:
        this.doc = this.app.loadDocument(template)

    filename      = os.path.split(template)[1]
    filename, ext = os.path.splitext(filename)
    filename      = '{}.odt'.format(filename)
    this.doc.saveAs(filename)

    return 0

if __name__ == '__main__':
   sys.exit(main())