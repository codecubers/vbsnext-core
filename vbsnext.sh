#!/bin/bash
set +x
WINEDEBUG=fixme-all wine cscript //nologo node_modules/vbsnext/bin/vbsnext.vbs "$@"

