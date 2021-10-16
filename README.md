# VB6NameSpaces (VB.dll) - BETA feedback wanted
https://www.youtube.com/watch?v=RJmDeGAroR4

![alt text](https://user-images.githubusercontent.com/39764372/127578868-cf190297-157c-46f1-a866-0e2e5721dcfd.png)

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

8-9-21
* Regfree does not work with 64 bit Office 2016 pro plus, displaying a "class not registered" error.

8-20-21
* Finally, I successfully made a standalone installer for VB6Porter with VB6Namespaces and VB65.  The installer is actually a minimal Nano install, plus a few extra files that were needed for stability and function, ie VB.dll (.net), msstdfmt.dll (link formatting), msaddndr.dll (addins), and VBA 6.5 core.  https://github.com/VBForumsCommunity/VB6Porter/raw/master/SetupVB6Porter.exe

8-23-21
* Installer now has an optional task to update files to the newest versions, this will replace any outdated files related to Visual Basic with VBA 6.5 (2007 or newer updates).  The installation is in-place, ie "Microsoft Visual Studio" folder inside the program files folder (x86).  The installer will now install msstdfmt.dll to SysWow64 directly, since innosetup kept mistakenly installing to System32.  Fixed, large projects on 64 bit systems no longer crash during compile/linking.

8-25-21
* The most important updated files to VB6 users are related to the compiler process, ie VB6DEBUG.DLL, LINK.EXE, and also MSPDB60.DLL are from 2007.

9-7-21
* Commandbar buttons initiate correctly now, using the IsDisposed property and isloaded flags.
* The installer includes an icon pack for explorer with icons from VBA 6.5 (2007).
* Icons are now restricted to only the ids available in the resource file, helps avoid overwriting icons that another program has modified, and also helps avoid incurring excessive Loading calls for a resource image that does not exist.
* Added the hidden easter egg with Visual Basic song:  "Show VB Credits" is included in the help menu.
* Added context menu items for semi-automatic code editing, ie Format procedure code, Format module code, and Format project code.  These functions will indent developer code in standard Visual Basic format.
* Form designer now saves contents to the class module without disturbing any declarations or events already generated by the developer.

9-21-21
* Changed names with underscore character to be displayed correctly despite the naming collision issue that was causing interop to hide names inside brackets, ie [Readonly].  The form designer may have some minor breaks to be resolved.  ...In the process of documenting functions inside vba6.dll to be used with the addin integration.  This dll exports a full pallette of internal low level functions that can be accessed for use.  At the core is Embedded Basic the Omega database engine that was retooled by Microsoft with intrinsic functions and UI.

10-2-21
* Fixed the form designer's feature to save to module.  Enumerations now retain the understore character to separate the property name from the property value.  Remove code by accident in previous update on 9/21/21.
* Discovered many new function declarations inside of the undocumented VBA6.dll.  A large portion of the exported function declarations can be called and detected by an intercept hook, without crashing the caller.  Everything going on behind the scences can be seen and understood without dissassembly.  Devolopers can duplicate and augment functionality to overcome the limitations of the extensibility model.
