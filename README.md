# python-mysql-excel

Example Python codes that do the processes between MySQL database and Excel spreadsheet files.

# Setup database table 

1. Use MySQLWorkbench and create container

![image](https://user-images.githubusercontent.com/78300596/140612534-04310ae6-d2a9-4df7-96f1-69dfdc0538f1.png)

# Install Python 3 and pipenv
1. Download Python 3 installation file from https://www.python.org/

2. Install pipenv as global package by this command.


```
pip install pipenv
```

Note: for macOS with pre-installed Python 2, use pip3 instead of pip or set version. Ex.pip3.9 install pipenv

# Install and run project by PyCharm
1. Download this project.

2. Open PyCharm and choose File -> Open... -> Then select project folder

3. ix any warning recommended by PyCharm.

4. Make sure that every project packages is installed, by open PyCharm Terminal and type command.
```
pipenv install
```

5. Open "import_data.py" or "export_data.py", right click on code area, select Run '{file name}'

