"""按姓氏批量发分入口。

源码运行时执行：

    python DM_surname.py --surname-batch

打包为名称包含“按姓氏”的 exe 时，会自动进入该模式。
"""

import sys

from DM import main


if __name__ == "__main__":
    if "--surname-batch" not in sys.argv:
        sys.argv.insert(1, "--surname-batch")
    sys.exit(main())
