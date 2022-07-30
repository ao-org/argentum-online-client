

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

