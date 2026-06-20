# Debate Keyword Finder
This is a macro you can add to your Microsoft Word for competitive debate to search for a set of key words in a document.

How might this be useful? You can run this macro to quickly search for offensive language or bad authors in your opponents' documents and punish them accordingly. You can also run this macro on your own document to make sure you don't do it too.

Portions of the creating the document and path-handling were adapted from [Shreeram Modi's](https://github.com/shreerammodi) own [debate scripts](https://github.com/shreerammodi/debate-scripts).

Email me at `blueabstractionoop@gmail.com` for any concerns.

## Installation
### 1. Downloading the Key Words Text File
Before you do anything in Word, you should download the `debate_words.txt` file. This is the file where you will store all the words you want the macro to search. I included some [commonly used words that have ableist connotations](https://web.augsburg.edu/english/writinglab/Avoiding_Ableist_Language.pdf) already, but you can definitely add your own.

Put the text file wherever is convenient for you. Just remember the path! You'll need it for later.

### 2. Opening Microsoft Visual Basic for Applications in Word
#### Mac
To open VBA for Mac, press `Option + F11`.

You can also use your mouse:

> Navigate to the menu bar > Tools > Macro > Visual Basic Editor

#### Windows
To open VBA for Windows, press `Alt + F11`.

You can also use your mouse:

> Navigate to the menu bar > File > Options > Customize Ribbon > Under Main Tabs, check Developer > OK

Then,

> Navigate to the menu bar > Developer > Visual Basic

### 3. Adding the Module for the Macro
You should now have the Visual Basic editor open.

For this installation, I'm assuming you have [Verbatim](https://paperlessdebate.com/) already. If you don't, download that first. 

Once you have Verbatim, expand `Debate` and right-click on `Modules`. Hover over `Insert` and click `Module` to create a new module. You can now paste everything from the `keyword_finder_macro.bas` file here. Now is when you need the path to your file of key words you want to search. At the top, underneath the green comment, change `YOUR\PATH\HERE\debate_words.txt` to your actual path.

## Setting Macros
Now you should have the code for the macro in Visual Basic. This step will show you how to set up the keys to run the macro.

If you're on Mac, go to:

> Tools > Customize Keyboard

Otherwise, if you're on Windows, go to:

> File > Options > Customize Ribbon > Customize

Change `Save changes in` to `Debate.dotm`.

After that, under Categories, click `Macros`. Then, search for `ScanDocument` under Commands. In the box where it says `Press new shortcut key`, you can pick a set of keys to press to activate the macro (I use `Command + ;` on Mac and `Ctrl + ;` on Windows, but you can pick anything else that's not already taken).

Press `Assign` and `Close`.

And you're done!

## Using the Macro
Using this macro is simple. Running this macro on a document will create a new copy of that document saved in your Downloads folder named "[K] document_name.docx" ("K" for key words). In that Word document, you will find every single time one of your key words is found in the `Comments` section.

Happy searching!
