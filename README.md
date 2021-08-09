# VB6NameSpaces (VB.dll) - BETA feedback wanted
https://www.youtube.com/watch?v=RJmDeGAroR4

This VB.NET (VB.dll) assembly makes it possible to fully interoperate with VBA/VB6 code.  Forms, component controls, and many namespaces are exposed to VBA/VB6 IDE.
The goal is to support VB.NET forms/controls/properties/events and Namespaces on VBA/VB6. 
BakcgroundWorkers; Multi-Threading; FileSystemWatchers; and 64 bit support...

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

7-10-21
* Images can be saved/loaded to/from the designer file.
* Constructors subs have been added to classes to match the same order as VB.NET contructors.
* Class buckets containing classes as properties can be used as an interface, ie when using the implements keyword in VB6

7-25-21 (internal version)
* Enumerations are saved/loaded to/from the designer file.
* DataGridView (Bound & Unbound) control added to toolbox
* DataSet (associated xml classes) component added 

7-27-21
* DataGridView now visible, optional hybrid control for vb6 toolbox

7-28-21
* XmlSqlClient classes added to compliment DataSet and DataGridView
* A few constructors were added and/or refined throughout the solution

7-29-21
* Test re-compile (internal version split) with Visual Studio 2022 preview 2.  Confirmed compatibility with .NET 4.8+ framework.

8-8-21
* A manifest file is now output for regfree com.  It seems to work here, but feedback wanted.
