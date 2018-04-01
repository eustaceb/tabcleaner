An extension for closing unused & non-project related documents. It works by crawling through all open projects in the solution and closing files that are not local to these directories. The LRU tab closing is not implemented yet.

The cleaning utility can be executed by right-clicking a tab and selecting "Close External Documents".

There is a single option under Tools->Options->TabCleaner, which lets the cleaner ignore modified (not saved) documents.
This is false by default, meaning the cleaner will not attempt to close such files. When set to true, a "Would you like to save these files" prompt will be displayed.