JEDI Code Library (modified for TRichView IDE installer)
=================

This version of JEDI Code Library is forked from https://github.com/project-jedi/jcl

A new branch "IDE Installer" is added, derived from version 2.6 (that was not the last version).

This modification is used to compile IDE Installer https://www.trichview.com/ideinstall/, an application for installing packages in Delphi and C++Builder IDE.

If you do not need other features of JCL, it's not necessary to compile any packages. Just add paths to "jcl\source\common", "jcl\source\include", "jcl\source\windows" to Delphi library paths.

You also have to download the jedi.inc and kylix.inc files from the https://github.com/project-jedi/jedi project and copy them to the "jcl\jcl\source\include\jedi" directory.
