# Python-misc
Python scripts to help easier several tasks.

MergExcel.py - Script for merging data from multiple excel documents to one excel document.
The excel docs must from the same template. This mean, its must have same column name, only different in the data. You got the idea.
You can merge all excel docs in the folder by specifying -d <directory-name>.
```
MergeExcel.py -d /path/to/directory/excels-files/ -o result.xlsx
```

You also can merge several excel files by using -f repeatedly.
```
MergeExcel.py -f test1.xlsx -f test2.xlsx -f test3.xlsx -o result.xlsx
```



