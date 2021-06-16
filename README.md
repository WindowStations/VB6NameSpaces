# VB6NameSpaces - Still in BETA
An single NET assembly now makes it possible to back-port VB.NET code to VB6 by interoperating with VB.NET UserControls and NameSpaces instanced as nested class buckets.
This is still a work in progress, however significant strides have been made to fully support VB.NET controls/properties/events and Namespaces directly through advanced dynamic interop. Including, but not limited to: BakcgroundWorkers; Multi-Threading; FileSystemWatchers; and 64 bit support are now made available to VB6.

* Step 1.  Download and unzip the VB project to the desired location.
* Step 2.  Install the VB.dll with SetupRegisterAssembly.exe (source included with drag/drop feature)
* Step 3.  Download and unzip hybridtest project.
* Step 4.  Test functionality and report issues here https://github.com/WindowStations/VB6NameSpaces/issues

Known issues:
* Control container ability is not included, ie DockStyle and Anchor properties are positioned relative to the UC, not the VB6 form.
* Images are not saved/compiled with the .frx file.
* Functions with overloads appear with _2 etc.  (updates planned)
* Opacity - VB.NET forms are now available as a full object, however the current control set is limited to common controls.  Controls do not have an opacity property yet.
