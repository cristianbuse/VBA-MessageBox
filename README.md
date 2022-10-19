# VBA-MessageBox
A custom message box that allows up to 5 buttons with custom labels.

![image](https://user-images.githubusercontent.com/23198997/196816620-2d98b7a3-edbb-4ded-a43a-26bd211041f0.png)

Extra abilities provided by ```MessageBox``` that the built-in ```MsgBox``` method does not provide:
1. Can display up to 5 buttons with custom labels and will return the label of the selected button
2. Displays a vertical scroll bar when the text is too large
3. The main text prompt can be copied from the locked MsForms.TextBox
4. The buttons get automatically resized and positioned based on how many buttons are displayed and how long their labels are 

Note that Cancel is allowed via the form's X button or via the Esc key if the form displays a single button or 'Cancel' is one of the button labels.

## Implementation
The class ```MessageBox``` creates an instance of the ```MessageForm``` userform and then adds all the necessary controls at runtime. It only exposes the ```Show``` method which allows the user to pass:
- the text for prompt
- the title
- the icon enum which corresponds to one of the 4 available icons
- the label(s) for up to 5 buttons
- the index of the default button

To remove the need for creating new instances, the class ```MessageBox``` has a global instance (```Attribute VB_PredeclaredId = True```). The exposed ```Show``` method is marked as a default member. It can be called like:  
```VBA
Debug.Print MessageBox("Test", "Title", icoInformation, "Button1", "Button2")
```
New instances of the ```MessageBox``` class can still be created but the approach used (global instance with default member) is more convenient.

The 4 icons (Critical, Exclamation, Information and Question) are the same as in the original ```MsgBox``` and they are saved in the ```Picture``` property of 4 labels situated on the ```MessageForm``` userform. The form only raises the ```QueryClose``` event. There is no other code within the form.

## Installation
Please avoid copy-pasting code. Either clone the repository or download the [zip](https://github.com/cristianbuse/VBA-MessageBox/archive/refs/heads/master.zip) and proceed importing the modules from there.

Just import the following code modules in your VBA Project:
* [MessageBox.cls](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageBox.cls)
* [MessageForm.frm](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageForm.frm) (you will also need the [MessageForm.frx](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/MessageForm.frx) when you import the frm)

## Demo
Import the following code module from the [demo folder](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/Demo) in your VBA Project:
* [Demo.bas](https://github.com/cristianbuse/VBA-MessageBox/blob/master/src/Demo/Demo.bas) - run ```DemoMain```

There is also a Demo Workbook available for download.

Screenshot examples:  
![image](https://user-images.githubusercontent.com/23198997/196816925-fc66a5f3-3b27-4d31-b321-8e4bacecb91d.png)

![image](https://user-images.githubusercontent.com/23198997/196816940-d8ab3984-ff39-4f50-bb45-79a541c4eb66.png)

## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
