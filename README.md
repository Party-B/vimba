# vimba
A MS Word template to give Vim-like functionality to your document.

# Notes
1. I developed this entirely within the Word VBA-IDE. So, my commits are gonna be random and likely just overwrite raw code (I'll try to keep it in two vba modules).
2. I used AI to generate some code, but it all sucked, so I just got some ideas from it. 
Problem was that it didn't really play well with legacy stuff
like VBA - which is what I taught myself to code on so I just found it to be easier to
do from scratch.
3. Run at your own risk - check code first to ensure you'd be happy with it. I didn't put anything intentionally volatile in, but don't trust me - my first commit comprised about 4 hours scrappy work.

# How to set up:
1. Create a new MS Word doc.dotm.
2. Alt + F11 to open the vba ide.
3. In the "ThisDocument" for the dotm file, paste the raw code from the commit.
4. Right click and create a new module named "vBinds".
5. In vBinds paste the vBinds raw code commit.
6. Set a Quick Access toolbar thing to launch the togglevimba macro.
7. NOTE!!! If you are doing this (Don't know how you found this code tbh), early breakages will mess with actual keybinds. Run the toggle macro a couple times to clear keybinds or create a new QAT toggle to specifically clear keybinds - note the TODO about it nuking all keybinds.

# TODO:
1.	Fix ClearBindings to only clear Vimba keys - Currently nukes all Word keybindings. Track in a collection?
2.	Unify buffer system architecture - Decide whether operators go through : buffer or have their own flow. Current d vs :w conflict will probably break things as more are added. Want to keep it vim like, so buffers for numbers and other actions.
4.	Implement ESC in buffer mode - Right now no way to cancel a buffer command.
5.	Fix the count/number system foundation - count prefix that ANY command can use. [affects every future motion/operator]
6.	Consolidate mode management - ThisDocument / vBinds Module for all mode logic. Split ownership is causing some headache.
7.	Add operator-motion grammar parser - Need logic to handle [count][operator][count][motion] pattern properly. Required before adding more operators.
8.	Add basic yank/put (y/p) - Core editing. Affects whether you need register system later.
9.	Fix case sensitivity bug - localKeyK types "K" instead of "k"
10.	Bind and implement a command - You have the function, just bind it.
11.	Complete basic motions - 0, $, ^, gg, G for line navigation
12.	Add insert mode variations - A, I, o, O for different insert positions
13.	Implement visual mode entry - v key, then make motions extend selection
14.	Add undo/redo - u and Ctrl+R bindings
15.	Implement line-wise operators - dd, yy, cc (operator pressed twice)
16.	Add search - / and n/N for navigation
17.	Implement text objects for operators - diw (delete inner word), da( (delete around parentheses). This is what makes Vim powerful, affects how you design operator logic.
18.	Add change operator - c to delete and enter insert mode
19.	Command history tracking - needed for next item.
20.	Implement dot command (.) - Repeat last action.


