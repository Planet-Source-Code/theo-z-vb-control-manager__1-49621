VB Control Manager 1.05
=======================

Introduction
------------
VB Control Manager is an ActiveX control to allow the user to dock controls and resize, move and show/hide them at run-time. It has 14 properties, 4 methods, and 18 events customly made, plus two collection: a Controls collection (with 20 properties) and a Splitters collection (with 27 properties). Open features.rtf file for the description of the VB Control Manager properties, methods, events and collections.

Basic Steps
-----------
To use the control, just place the VB Control Manager control on your form and add several controls on it. That's it! You don't have to add any codes to your form to have basic features of VB Control Manager.

And here's what you'll get:
* all controls on the VB Control Manager control are docked and filled the control
* all splitters needed to resize the controls at run-time have been built
* all title bars needed to move the controls at run-time have been built
* all close buttons on the title bars needed to close the controls at run-time have been built
* the VB Control Manager control will automatically resize itself to fill its container everytime its container is resized
 
What's New
----------
VB Control Manager is the improvement of VB Splitter, an ActiveX control to resize docked controls at run-time. I change the name because VB Control Manager can do a lot of more than just to split/resize control.

Here's "what's new" in VB Control Manager compare to VB Splitter:
* Added more than 3500 lines of code
* Added move control features
* Added show and hide control features
* Automatically call Activate method when the control is activated (developers now don't have to add any codes in their from to make VB Control Manager activated)
* Added 3 properties: TitleBar_CloseVisible, TitleBar_Height and TitleBar_Visible
* Added 3 methods: CloseControl, MoveControl and OpenControl
* Added 11 events: ControlAfterClose, ControlBeforeClose, ControlMove, ControlMoveBegin, ControlMoveEnd, SplitterMoveBegin, TitleBarClick, TitleBarDblClick, TitleBarMouseDown, TitleBarMouseMove and TitleBarMouseUp
* Made Controls collection and Splitters collection become public
* Added 5 properties to Controls collection: Closed, Name, TitleBar_CloseVIsible, TitleBar_height and TitleBar_Visible
* Added 5 error handlers: Error on resizing a splitter, error on moving a control because of there is not enough room to drop the control, invalid control id, invalid splitter id and error on moving a control because of the control is closed

More Notes
----------
* This ActiveX control uses subclassing, so DO NOT stop the debugger with the IDE STOP BUTTON or else VB may crash
* To have the controls docked at design-time, simply resize the VB Control Manager control
* To have the resize, move and show/hide features activated at design-time, right-click the VB Control Manager control and select Edit from the context menu

Credits
-------
* To Steve McMahon for the SubTimer6.vbp and for the excellent articles at http://www.vbaccelerator.com 
* To KPD-Team at http://www.allapi.net for their API-Guide
* To Carles P.V. (carles_pv@terra.es), Vlad Vissoultchev (wqw@bora.exco.net), Paul Caton (paul_caton@hotmail.com) and umairata (umairata@hotmail.com) for various bugs information

Use and Distribution
--------------------
* This control is freeware and opensource
* You may freely modify the source only to be used on your own program
* You may NOT modify the source and/or publish it as your work without permission from the author
* The author provides the control "as is" with no warranty of any kind

The author
----------
Theo Zacharias (theo_yz@yahoo.com)