# ppt-font-change-macro

## What is this (short)?
Macro for changing fonts for all slides in MS Office Powerpoint.

### What does this do (longer explaination)?
This macro lets you manually enter the font name and that font is applied to all the text across all the slides in the presentation. 

### What is it good for? 
This can be useful when you are reformatting a large file or for translation etc. 
In general, whenever you have to change the font for a lot of slides but don't want to do it manually. 
You can use it for study too. The fact that Microsoft does not have any batch processing features for formatting is a bit concerning. :S 

## What does it work on?
It works for text in:
- Slides
- Shapes
  - Sub-shapes/Grouped shapes
- Tables.
- Any combination of the above

It does not work for
- Embedded charts/excel data.

## How does it work?
The implementation uses simple for-loops to iterate through all the slides, shapes, sub-shapes etc. 
This is perhaps not the best implementation, but this is what I came up with after not finding anything relevant to what I wanted. 
I am not familiar with writing macros or visual bas

## Why does this exist?
I and a few others often had the need to convert text from one font to another.
For example, when translating something or when changing formatting of text. 
I am posting this here because: (a) I no longer have a clause stating that any code I created is the property of my company, and (b) hopefully more people with similar problems can find this and use it. 
