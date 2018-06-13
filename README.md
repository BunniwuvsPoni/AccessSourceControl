# AccessSourceControl
VBS script to provide version control for Access applications.

This was created using the source: https://stackoverflow.com/questions/187506/how-do-you-use-version-control-with-access-development

How to use:
1. Edit Source.bat and Rebuild.bat files to reflect your Access application name. Optional: Add a <path> after the application name to determine where the files should go, otherwise the default path is "\Source"
2. Run Source.bat to convert all modules, classes, forms, queries and macros from an Access file to text and saves the results in separate files to <path>.
3. Run Rebuild.bat to recompose your Access application using the text files in <path>. NOTE: This will overwrite all modules, classes, forms, queries and macros.
