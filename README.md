one-liner helper to read Excel XLSX
====================

Usage
--------------------

```sh
% XLSX_ARGV.pm mybook.xlsx
mySheet1
% perl -MXLSX_ARGV=mybook.xlsx -nle 'print if m{^sheetData} .. m{^/sheetData}' mySheet1
sheetData
row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="1"
c r="A1" s="1" t="s"
v>0</v
/c
c r="B1" s="1" t="s"
v>1</v
/c
c r="C1" s="1" t="s"
v>2</v
/c
/row
row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="2"
c r="A2" s="0" t="n"
v>1</v
/c
c r="B2" s="1" t="s"
v>3</v
/c
c r="C2" s="1" t="s"
v>4</v
/c
/row
/sheetData
```

Basic idea
--------------------

Writing a code to extract data from xml in general is cumbersome task.
It is boring and can be time consuming. But if we focus on particular context,
we might be able to change it like an ordinary text filtering one-liner.

Followings are key ideas: 

* Use ``$/ = "><"`` in Perl.
* Tie ``@ARGV`` to make ``while (<>)`` loop into selected sheets.
* Provide sheetName access for command line.
* Make the module itself runnable.
