#!/bin/bash
set +x
WINEDEBUG=fixme-all wine cscript //nologo node_modules@vbsnext/vbsnext-core/bin/vbsnext.vbs "$@"

