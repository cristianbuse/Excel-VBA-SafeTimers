## Update Feb 2023

This solution is no longer maintained here or on Code Review, mainly because I do not need a timer solution anymore. The only purpose of this approach was to get a reliable, crash-free call that can get Excel out of UDF mode - which I've eventually achieved by posting a ```WM_DESTROY``` message to a userform. See [here](https://github.com/cristianbuse/VBA-FastExcelUDFs). This repository is now purely an exercise in how to work from a remote application instance and get somewhat reliable timer functionality without crashes. **I do not recommend using this solution into an actual VBA project!**

***

# Excel-VBA-SafeTimers
Safe Windows API timers for Excel.

Related [Core Review question](https://codereview.stackexchange.com/questions/274652/safe-windows-api-timers-for-excel)

In VBA (Excel) there is no reliable timer available. The ```Application.OnTime```method fails if called from a UDF (User Defined Function) context or while debugging code. The other most used solution is using Windows API timers but those are notorious for crashing.

This repository contains 2 different solutions to the same problem. Both solutions work based on the same principle. A second instance of Excel is created and code is added to that instance. Note that **there is no need to have trusted access to the VB Object Model**!

Solutions:
 1. A native solution that works on both Windows and Mac - calls ```Application.Run``` via a 'callback' object.
 2. A Windows API timers solution that minimizes the risk of application crashes - [posts](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-postmessagea) messages back to the relevant application instance.

**If multiple books within the current application instance are using timers, then the same remote app instance will be used for all of them**. The remote code is 'smart' enough to manage multiple workbooks.

The remote application is automatically terminated when the current application is closed.

Quick comparison of both solutions:

|#|Approch|Platform     |Crashes|App busy|Debugging|Formula Edit|Modal app dialog|Modal Userform|
|-|-------|-------------|-------|--------|---------|------------|----------------|--------------|
|1|Native |Windows + Mac|None   |Waits   |Waits    |Waits*      |Waits           |Executes      |
|2|Win API|Windows Only |Rare   |Waits   |Waits    |Executes    |Executes        |Waits         |

<sup>*editing formulas is slowed down while there are timers pending. If no timer is pending then speed is normal.</sup>

## Native Solution #1

A new ```TimerCallback``` instance is created for each timer and passed to the remote application. The remote application calls the ```TimerProc``` method of the ```TimerCallback``` class which in turn uses ```Application.Run``` to call a macro. This implies that when creating a new timer the name of the method must be passed as a text.

When creating a new timer using the ```CreateTimer``` method, there are 2 options for the delay:
 1. delay is set to 0 - the callback will only be called once but the call is guaranteed to happen.
 2. delay is > 0 - the macro will be called repeatedly in intervals approximately equal to the provided delay and will only stop if the timer is removed or state is lost.

### Installation
Just import the following code modules in your VBA Project:
* [**LibTimers.bas**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%201%20-%20Native%20(Win%20%2B%20Mac)/LibTimers.bas)
* [**TimerCallback.cls**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%201%20-%20Native%20(Win%20%2B%20Mac)/TimerCallback.cls)

### Demo
Import the following code module:
* [Demo.bas](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%201%20-%20Native%20(Win%20%2B%20Mac)/Demo/Demo.bas) - run ```DemoMain```

## Windows API Solution #2

The ```LibTimers``` module exposes wrappers for ```SetTimer``` and ```KillTimer``` which match the exact function signatures as the Windows counterparts. This makes it easy to update existing projects to use this library.
There is only one 'real' main timer per workbook and it's the only one affected by the posted messages from the remote app. When called, it safely [dispatches](https://docs.microsoft.com/en-us/windows/win32/api/winuser/nf-winuser-dispatchmessage) messages to the relevant procedures (timer procs).
No timers are left hanging. Even if state is lost, the remote app will make sure to call the book timer so that it can remove itself.

### Installation
Just import the following code modules in your VBA Project:
* [**LibTimers.bas**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%202%20-%20Windows%20only%20APIs/LibTimers.bas)
* [**SafeDispatch.cls**](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%202%20-%20Windows%20only%20APIs/SafeDispatch.cls)

### Demo
Import the following code module:
* [Demo.bas](https://github.com/cristianbuse/Excel-VBA-SafeTimers/blob/master/src/Solution%202%20-%20Windows%20only%20APIs/Demo/Demo.bas) - run ```DemoMain```

## License
MIT License

Copyright (c) 2022 Ion Cristian Buse

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
