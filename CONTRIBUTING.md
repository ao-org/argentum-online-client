

# Coding style

## Do not use uppercase letters
Please when submitting new pull requests do not use uppercase letters in the code.

For example, instead of `ComputeSomethingNice` use `compute_something_nice`

Always use lowercase letters and underscore.

## Always use either `byref` or `byval`

`Sub compute_something_nice(byval arg1, byval arg2)`

## Do not include your name in the code

Please do not include comments with your name and date among the code as shown below because this information can easily be found using git.
```
   ' Author: Your Name 
   ' Last Modify Date: XX/ZZ/YYYY
```
## Run `git_ignore_case.sh`
Unfortunately VB6 IDE messes up the code automatically turning some lower case letters to upper case for some identifiers in the code.

For example VB6 will turn 
`Call WritePreguntaBox(Destino, UserList(Origen).Name & " desea comerciar contigo. ¿Aceptás?")` into
`Call WritePreguntaBox(Destino, UserList(Origen).name & " desea comerciar contigo. ¿Aceptás?")`

- Please run the script `git_ignore_case.sh` before committing any changes to a file
- Make sure there are no changes in existing code which only rename identifier as described above
- Intential renaming of existing identifiers (functions,subs, variables) is highly encouraged. Any new names must be in english and use only lowercase and _ as in `new_name`
