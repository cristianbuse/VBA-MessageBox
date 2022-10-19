# VBA-MessageBox
A custom message box that allows any label on up to 5 buttons

Extra abilities provided by ```MessageBox``` that the built-in ```MsgBox``` method does not provide:
1. Can display up to 5 buttons with custom labels and will return the label of the selected button
2. Displays a vertical scroll bar when the text is too large
3. The main text prompt can be copied from the locked MsForms.TextBox
4. The buttons get automatically resized and positioned based on how many buttons are displayed and how long their labels are 

Note that Cancel is allowed via the form's X button or via the Esc key if the form displays a single button or 'Cancel' is one of the button labels.

## Implementation
The class ```MessageBox``` has a global instance (```Attribute VB_PredeclaredId = True```) and a default member (```Show```) so that it can be called like:  
```VBA
Debug.Print MessageBox("Test", "Title", icoInformation, "Button1", "Button2")
```

## Use

## Installation
Please avoid copy-pasting code. Either clone the repository or download the [zip](https://github.com/cristianbuse/VBA-MessageBox/archive/refs/heads/master.zip) and proceed importing the modules from there.

Just import the following code modules in your VBA Project:
* [MessageBox.cls](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageBox.cls)
* [MessageForm.frm](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageForm.frm) (you will also need the [MessageForm.frx](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageForm.frx) when you import the frm)

## Demo
Import the following code module from the [demo folder](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/Demo) in your VBA Project:
* [Demo.bas](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/Demo/Demo.bas) - run ```DemoMain```

There is also a Demo Workbook available for download.

## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.