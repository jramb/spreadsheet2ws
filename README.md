# Spreadsheet to WS

## Intro

This is a little tool which allows you to call a SOAP web service repeatedly using the contents of a Spreadsheet file (spreadsheet-format)
as input.

Any column in the spreadsheet can be used for creating the request XML.
When the tool is executed it will go through any row in the spreadsheet file (first tab only)
and (using a series of simple, customizable stylesheets) convert the line into a request.
The web service is called and the result is stored in the same spreadsheet file again.
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

Configuration and parameters: Parameters can be used in several places in the form of Key Values.
First of all, all parameters are loaded from `setup\config.xml`.

You can complete, add and override any parameter in `config.xml` using the **second** tab in the spreadsheet file.
The column A and B can contain and override any parameter. (If you use a header row, just leave A1 empty to not risk any accident.) Only the first two columns are used as-is, the rest can be used for comments.
Lines which have no value in A are not used.
Parameters specified here override `config.xml`.

In the last step you can override any parameter in the run command line, as parameters in the form `key=value`. These are strongest and override both parameters in the file and in `config.xml`.

Why is this so a big deal? Because the parameters (all of them!) can be used in the stylesheets as input, which is quite powerfull.

* Prepare the .xlsx file with the data you want to load. In the current setup, only the **first** tab and there only the columns where the first column (`A` = `0`) is empty or contains the text `NEW` is loaded.
* Close(!) the file in the spreadsheet program. Spreadsheet program lock the file while open and the our program need to update the file (unless `update-file` is set to false!).
* Make sure that java is in your path, in `cmd` shell go to the directory which contains the `run.bat` file, the `lib/` and the `setup/`.
* Run it:
    `run`

* The first parameter (and only the first) is checked if it is the name of a readable file. If yes, that file is the spreadsheet file to be loaded instead of what might be specified in the `config.xml`.
* As mentioned you can override any parameter in the config.xml in the `run` command, e.g. to prevent update of the spreadsheet file:
    `run update-file=false`
or if you want to run in a different environment:
    `run environment=TEST`
or you can add additional parameters (which can be used in the stylesheets):
    `run keyword=value01 debug=false password=secretword username=admin`

Unless run with `update-file=false` the spreadsheet file is updated according to your setup after the run is complete.


## Warning

Because the way spreadsheet data works, this tool will not use or evaluate formulas,
sorry. You will need to evaluate formulas yourself and paste the value in the correct columns.
I do not see an easy way around this.

Last thing: I take no responsibility of what happens if you use this tool.
If it somehow harms your data: your problem, not mine.

## Author

Jörg Ramb


## TODO

* Do we need parallel execution?

