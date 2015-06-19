# TableSheet

Конвертор таблиц в стандартный CSV формат



Installation
-------------


### requirements.txt

    chardet==2.3.0
    xlrd==0.9.3
    xlsx2csv==0.7.1
    xlutils==1.7.1
    xlwt==0.7.5



## Usage 



Преобразовать в правельный CSV формат, разделитель ```,```

    $ python pyfcsvconv.py file.xls
    $ python pyfcsvconv.py -d ';' file.xls
    $ python pyfcsvconv.py -i f2.csv


