# Excel-VBA-SafeTimers
Safe Windows API timers for Excel.

Windows API timers are notorious for crashing. This project is an attempt to minimize the risk of application crashes.

A second instance of Excel is created and code is added to that instance. Note that **there is no need to have trusted access to the VB Object Model**! That second instance runs a continuos loop which [posts](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea) messages back to the application instance using this project.

The ```LibTimers``` module exposes wrappers for ```SetTimer``` and ```KillTimer``` which match the exact function signatures as the Windows counterparts. This makes it easy to update existing projects to use this library.
There is only one 'real' main timer per workbook and it's the only one affected by the posted messages from the remote app. When called, it safely [dispatches](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-dispatchmessage) messages to the relevant procedures (timer procs).
No timers are left hanging. Even if state is lost, the remote app will make sure to call the book timer so that it can remove itself.

If multiple books within the current application instance are using timers, then the same remote app instance will be used for all of them. The remote code is 'smart' enough to manage multiple workbooks.

The remote application is automatically terminated when there are no more books left to manage i.e. the current application was closed.

Related [Core Review question](https://codereview.stackexchange.com/questions/274652/safe-windows-api-timers-for-excel)

## Installation
Just import the following code modules in your VBA Project:
* [**LibTimers.bas**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/LibTimers.bas)
* [**SafeDispatch.cls**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/SafeDispatch.cls)

## Demo
Import the following code module:
* [Demo.bas](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Demo/Demo.bas) - run ```DemoMain```

## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
