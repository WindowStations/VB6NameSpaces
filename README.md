# VB6NameSpaces
For back-porting .NET samples to VB6 by extending VB6 with .NET duplicates in nested classes like namespaces


Nested classes in VB6, a house of bricks.
Although the class builder in VB6 isn't stable enough to work with, so I recommend initializing your classes manually.

This is a work in progress, so some areas may be underdeveloped or blank, while others are a little more refined.
For example, see the classes for Process, File, Directory, MessageBox, and Screen.
The Screen class required a Rectangle class and a Point class to be made, like the .NET class counterpart.

More imports are planned of course. 
Please contribute if you have any good API samples that will help duplicate .NET imports/namespaces.

Alpha preview: https://vb6x.org/

Advantages
1. Compatibility. I would like to facilitate easier conversion/s between VB.NET and VB6.
The goal is to add unique code resources and concepts that have been developed in .NET over the years

2. Speed. Convert some net desktop code that is best suited for quick/direct desktop operations using the API.

3. Longevity. Maintain compatibility regardless of the net framework version installed, now and in the future.

4. Flexibility. New classes and/or features can be added or removed easily depending on your expertise.
Different templates can be made to group up classes as needed.

5. Robustness. Murphy's law applies to Visual Studio and the net framework.
Visual Studio NET also requires a full installation of a very particular framework to target with.
