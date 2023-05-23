# PyExcelSync

PyExcelSync is a tool to get data from another excel file and write the value to existing file.
<br /> <br />
**Example** <br />
Take a look at **SourceName**
<br /> <br />
In this case, we need to get the value (**NO SURAT**) of each file and append on the column.
<br />
But the problem is... it will take a lot of times if we open one by one right? <br />
Then imagine there is a lot of source file 😭
![This is a alt text.](https://i.imgur.com/ZbLXPfC.png "This is a sample image.")
![This is a alt text.](https://i.imgur.com/Ba6nbVV.png "This is a sample image.")
That's why we need a tool to make our work easier! <br />
Using **PyExcelSync** will automatically read the values of each files,
![This is a alt text.](https://i.imgur.com/4iqqxVq.png "This is a sample image.")
Then it will add the result to new column, next to **SourceName**
![This is a alt text.](https://i.imgur.com/IUNW24s.png "This is a sample image.")

## Installation

**PyExcelSync** requires [Python3 & Pip3](https://www.python.org/downloads/) to run.

Create **virtualenv** first,
Then install **requirements.txt** for the dependencies.

Skip this if you already install the **virtualenv**

```sh
pip3 install virtualenv
```

Create **virtualenv** for this project

```sh
virtualenv -p python3 <name>
```

Activate the **virtualenv**

```sh
source <name>/bin/activate
----
to deactivate
deactivate
```

Install **requirements.txt** for the dependencies.

```sh
pip install -r requirements.txt
```

## Running

**Make sure you read this step carefuly before running**

- The default file name for work file is **works.xlsx**

```sh
Just replace the file if you want to use ur own with the same name (works.xlsx)

Or change this line to ur own filename (main.py line 19)
file_work = "works.xlsx"

Make sure the extension file is .xlsx!
```

- The default file header is **NoSurat**

```
Or change this line to make ur own header name (main.py line 23)
dataSurat.append('NoSurat')

Make sure the header doesn't have a space!
```

**The most important** <br />
Make sure you put the source file in list **SourceName** at **file/** folder.

**Running the tools**
In your terminal (virtualenv)

```sh
python main.py
```

And check the result in **output.xlsx** file
