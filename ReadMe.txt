' =========================================================================================
' 3D Computer Graphics for Visual Basic Programmers: Theory, Practice, Source Code and Fun!
' Version: 6.0 beta - Precision Edition
'
' by Peter Wilson
' Copyright © 2004 - Peter Wilson - All rights reserved.
' http://dev.midar.com/
' =========================================================================================

Overview
========
Learn how to create your own 3D graphics library without using OpenGL or DirectX. This 3D application allows you to load DirectX data files into the application and view them in 3D. The virtual camera is cantered on the origin (although you can change this). Use the mouse to move the virtual camera around the object. This 3D application uses a Right-Handed Coordinate system (as opposed to DirectX that uses a left-handed coordinate system). The 3D maths is based on industry recognized standards as found in 'Computer Graphics Principles and Practice, Foley*vanDam*Feiner*Hughes'. The 3D maths uses column-vector notation and is very stable; it won't change too much in the future, so this would be a good project to get familiar with. This is a solid project that I will be improving on. I've listed the code as intermediate as I don't think its too hard to follow. Comments are everywhere. For those of you familiar with my previous works (look them up), this one does not have any fancy music or animation as it's focus is somewhat more serious.


Instructions
============
1) Reset the Application using the "Reset" menu.
2) Import a *.x data file (supplied)
3) Use the Mouse in combination with the Left/Right mouse buttons to
   move the position of the virtual camera. The camera coordinates will
   be displayed in the titlebar of the window.
4) When you want to view another 3D file, go back to step 1.


Problems / Limitations
======================
* I'm still cleaning up the DirectX import routine (although it is fairly stable). As a result, if you use your own *.x files they may need some massaging to import correctly. This is somewhat normal with DirectX files.


Send me an e-mail if you need help with anything.

Enjoy!

Peter Wilson
peter@midar.com
http://dev.midar.com/
