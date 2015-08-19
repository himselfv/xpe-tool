# xpe-tool #
This command-line tool allows you to simplify some tasks you might encounter while designing your Windows XP Embedded image.

### Functions ###
Find all revisions of a component:
```
xpe find "component name"
```
Use % as a replacement for any set of symbols.

Display component info:
```
xpe info <component-id>
```

Print component dependencies:
```
xpe deps <component-id>
```

Export list of files contained in a component or components:
```
xpe files <component-id> [component-id] ...
```

Collect all the files contained in a component:
```
xpe collect-files <component-id> <target-dir>
```

Create a registry export file (.reg) for the component from its "Registry" resources:
```
xpe registry-export <component-id> [component-id] ... >myfilename.reg
```

### Setup ###
Download the code and compile it against any version of Delphi, or download the compiled version from Downloads.

Create ADO connection to your MSSQL server containing XP Embedded database. Specify ADO connection parameters in "xpe.cfg".

Use the tool. If you want it to be accessible from any folder, add it to the PATH environment variable.