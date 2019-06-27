# `xtpxlib-TBD`: Xatapult XML Library - TBD

Version release and dependency information: See `/version.xml` 

Xatapult Content Engineering - http://www.xatapult.nl

Erik Siegel - erik@xatapult.nl - +31 6 53260792

----

**`xtpxlib`** is a library containing software for processing XML, using languages like 
XSLT, XProc etc. It consists of several separate components, named `xtpxlib-*`. Everything can be found on GitHub ([https://github.com/eriksiegel](https://github.com/eriksiegel)).

**`xtpxlib-TBD`** ([https://github.com/eriksiegel/xtpxlib-TBD](https://github.com/eriksiegel/xtpxlib-TBD)) is TBD.

----

## Using `xtpxlib`

* Clone the GitHub repository to some appropriate location on disk. That's basicly it for installation.
* If you use more than one `xtpxlib` component, all repositories must be cloned in the same base directory.

----

## Library contents

### Directories at root level

| Directory | Description | Remarks |
| --------- | ----------- | --------|
| `data` | Static data files. |  |
| `etc` | Other files, mostly for use inside oXygen. |  |
| `xplmod` | XProc 1.0 libraries. |  |
| `xsl` | XSLT scripts. |  |
| `xslmod` | XSLT libraries. |  |

The subdirectories named `tmp` and  `build` may appear while running parts of the library. These directories are for temporary and build results. Git will ignore them because of the `.gitignore` settings.

Most files contain a header comment about what they're for.
