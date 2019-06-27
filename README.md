# `xtpxlib-xoffice`: Xatapult XML Library - MS Office file handling

Version, release and dependency information: See `/version.xml` 

Xatapult Content Engineering - http://www.xatapult.nl

Erik Siegel - erik@xatapult.nl - +31 6 53260792

----

**`xtpxlib`** is a library containing software for processing XML, using languages like 
XSLT, XProc etc. It consists of several separate components, named `xtpxlib-*`. Everything can be found on GitHub ([https://github.com/eriksiegel](https://github.com/eriksiegel)).

**`xtpxlib-xoffice`** ([https://github.com/eriksiegel/xtpxlib-xoffice](https://github.com/eriksiegel/xtpxlib-xoffice)) is TBD.

----

## Using `xtpxlib`

* Clone the GitHub repository to some appropriate location on disk. That's basicly it for installation.
* If you use more than one `xtpxlib` component, all repositories must be cloned in the same base directory.

----

## MS Office file handling

MS Office files (Word, Excel, etc.) are zip files that contain a lot of XML and other files. The format is very complicated and rather over-engineered. This component contains some XProc (1.0) libraries to handle MS Office files. Current functionality is:

* MS Word (`.docx`):
  * Convert to XML format
  * Convert from this XML format back to Word (using a template Word file)
* MS Excel (`.xlsx`):
  * Convert to XML format
  
The Word XML format is not very sophisticated (yet). Maybe when the need arises this will be changed and made more intelligent.

The Excel format contains a lot of information and is used quite a lot in several applications. There is even a schema for this format in `/xsd/xlsx-extract.xsd`.

----

## Library contents

### Directories at root level

| Directory | Description | Remarks |
| --------- | ----------- | --------|
| `etc` | Other files, mostly for use inside oXygen. |  |
| `xplmod` | XProc 1.0 libraries for converting MS Office files. |  |
| `xsd` | XML schemas. |  |
| `xslmod` | XSLT libraries. | These libraries were designed for internal use, but might be interesting if you are navigating around in an MS Office file (in `xtpxlib-container` format) someplace else.  |

The subdirectories named `tmp` and  `build` may appear while running parts of the library. These directories are for temporary and build results. Git will ignore them because of the `.gitignore` settings.

Most files contain a header comment about what they're for.
