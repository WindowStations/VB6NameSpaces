# VB6NameSpaces - BETA feedback wanted
A single NET assembly now makes it possible to back-port VB.NET code to VBA/VB6 by interoperating with VB.NET UserControls and NameSpaces instanced as nested class buckets.
This is still a work in progress, however significant strides have been made to fully support VB.NET controls/properties/events and Namespaces directly through advanced dynamic interop. Including, but not limited to: BakcgroundWorkers; Multi-Threading; FileSystemWatchers; and 64 bit support are now made available to VBA/VB6.

* Step 1.  Download and unzip the VB project to the desired location.
* Step 2.  Install the VB.dll with SetupRegisterAssembly.exe (source included with drag/drop feature)
* Step 3.  Download and unzip hybridtest project.
* Step 4.  Test functionality and report issues here https://github.com/WindowStations/VB6NameSpaces/issues

Known issues:
* Hybrid (VB.NET controls on VB6 forms) Control container ability is not included, ie DockStyle and Anchor properties are positioned relative to the UC, not the VB6 form.
* Hyrbrid control Images are not saved/compiled with the .frx file.
* Functions with overloads appear with _2 etc.  (updates planned)
* Hybrid opacity is not currently supported.
* Opacity - VB.NET forms are now available as a full object with a full set of common controls.
* VBA object property grid values/names are mistmatched.   The code broke with the latest set of property updates, so it will need to be dialed in again, once the main interop is complete

Progress updates:
* Error warnings were resovled for obsolete functions.  Class errors due to missing/wrong-type properties/params were hooked up.
* A new VB.NET form/control designer for VB6 allows developers to design VB.NET forms and controls with a modern interface.  Thus avoiding the limitation of the VBA/VB6 property windows and the limitations of hybrid controls.  This designer is now the primary focus under development, for stand-alone and integration with VB6Porter.dll.
* Refinements have been made to begin the completion of the code base.  Individual constructor subs will have the same parameter order as their VB.NET counterparts.
* The VB Form Designer can reload the new designer format for VB6.
* Several new classes have been added.  See the class diagram for any new additions to the VB.NET/VB6 framework: https://github.com/WindowStations/VB6NameSpaces/issues/1
