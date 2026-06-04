# -*- coding: utf-8 -*-
paths = (
    "Documents/А_ОтражениеЗПпоКазне/Ext/ObjectModule.bsl\n"
    "AccumulationRegisters/А_ВзаиморасчетыССотрудниками/Ext/ManagerModule.bsl\n"
)
out = r"C:\Configuration_downloads\BASERP25\_Rarzrabotki\notebook\knowledge_А_ОтражениеЗПпоКазне\Python\_artifacts\loadlist.txt"
open(out, "wb").write(b"\xef\xbb\xbf" + paths.encode("utf-8"))
print("listfile created:", out)
