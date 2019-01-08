# Error-Handling-and-Debug-Prints
## MS Access Tools, Standard Error Handler and Debug Print Procedures

This file contains a single VBA module. It contains the procedures that I use to do standard run-time error handling and a few rather simple tools to log and print variables for debugging. These procedures do use Access Objects but Late Binding is used so that References need not be set to use them.

The code is, I think pretty well documented in the comments.

These are the main procedures:
1. ErrHandler - displays run-time errors and logs them to a file on c:\
1. Say - for debugging; display a scalar or string
1. WriteLog - for debugging; writes message to log file
1. Trace - for debugging; writes message tagged with current procedure to log file
1. ProcTemplate - commented example of the procedure template with error handling that I use
