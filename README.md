# Spreadsheet to WS

## Intro

This is a little tool which allows you to call a SOAP web service repeatedly using the contents of a Spreadsheet file (Excel-format)
as input.

Any column in the spreadsheet can be used for creating the request XML.
When the tool is executed it will go through any row in the spreadsheet file (first tab only)
and (using a series of simple, customizable stylesheets) convert the line into a request.
The web service is called and the result is stored in the same Excel file again.
You can also update any column from the result of the call, which allows you to mass-query data at the same time.

It is a very simple tool to use, but very versatile. Since anything is configured using stylesheets and an XML config.xml
(where every parameter can be overridden during execution!), you can do quite a lot of different things with it.

I used this for two of my customers already. Since it it so simple (but useful, I think), it would be a shame not to
offer it to others to reuse. Need to work more on some documentation though.

One wheel less to invent, I hope.

# How to use

## Preparation

* Check out the project
* add the necessary library jars from the Apache POI and XMLBeans project (see file under `libs/`)
These external libraries are needed:
https://poi.apache.org/, https://xmlbeans.apache.org/ 

* poi-3.11-20141221.jar
* poi-ooxml-3.11-20141221.jar
* poi-ooxml-schemas-3.11-20141221.jar
* xmlbeans-2.6.0.jar

(Last tested with these versions)

* Modify the files under `setup/` according to what you need. This is a working example, but you need to modify it.

## Run

* Prepare the .xlsx file with the data you want to load. In the current setup, only columns where the first column (`A` = `0`)
is empty or contains the text `NEW` is loaded.
* Close(!) the file in Excel. Excel locks the file and the command needs to update the file.
* Make sure that java is in your path, in `cmd` shell go to the directory which contains the `run.bat` file, the `lib/` and the `setup/`.
* Run it:
    `run`

* You can override any parameter in the config.xml in the `run` command, e.g. to prevent update of the Excel file:
    `run update-file=false`
or if you want to run in a different environment:
    `run environment=TEST`
or you can add additional parameters (which can be used in the stylesheets):
    `run keyword=value01 debug=false password=secretword username=admin`

Unless run with `update-file=false` the Excel file is updated according to your setup after the run is complete.

## Autor

Jörg Ramb


## TODO

* Make this a Maven project?
* Do we need parallel execution?
