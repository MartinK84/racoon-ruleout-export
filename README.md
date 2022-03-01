# RACOON Ruleout Export
A simple python script for parsing the xml-file of a Mint RACOON export and anonymization of the data using a sha256 hash. The script will convert the given XML file to a much smaller Excel file and include only those columns that are required for the Ruleout training.

## Run the pre-compiled binary ruleout_export.exe
```ruleout_export.exe``` is compiled using pyinstaller on windows 10 anaconda environment. To run the program simply execute ```ruleout_export.exe``` and select the previously exported Mint xml-file in the file selection popup. The parsed and converted excel file will be stored at the same location/path as the input file. 

## Run the script using Python
```python ruleout_export.py -i input.xml -o output.xml -v```

## Notes
The source code is based in parts on work by Moon Kim and the RACOON xml parser at https://gitlab.com/moon.kim.mail/racoon-xmlparser

## License

This script is provided under the terms of the MIT License (MIT)

Copyright (c) 2010 New York University

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
