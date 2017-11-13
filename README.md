# TranslatePPTX

<span STYLE="font-variant: small-caps;">translatepptx</span>
is a very simple code based on the
[Apache POI](https://poi.apache.org)
for facilitating editing or translation of Powerpoint
presentations in `.pptx` format.

Using
<span STYLE="font-variant: small-caps;">translatepptx</span>
to edit or translate a Powerpoint file is a three-step process:

1. You do a first run of
   <span STYLE="font-variant: small-caps;">translatepptx</span>
   on your `.pptx` file in *text extraction* mode.
   <span STYLE="font-variant: small-caps;">translatepptx</span>
   extracts all text strings on all slides in the presentation 
   (including all text boxes, all text in graphs, etc.) and 
   writes these out, with some identifying labeling info,
   to a **plain text file**.

2. You edit this file in your favorite text editor,
   modifying any text strings that you want to edit or translate
   and deleting all the rest. 

3.  Finally, you pass your edited list of text strings as a
    command-line argument to a second run of
    <span STYLE="font-variant: small-caps;">translatepptx</span>,
    now running in *text replacement* mode;
    this time the code replaces all the strings you edited with 
    your edited versions and writes out a new `.pptx` file
    reflecting your edits.

## A tutorial example

Here's an example involving a presentation called
[`MyDeck.pptx`](example/MyDeck.pptx),
which contains two slides and has a couple of text boxes
and a graph:

### Step 1: Extract text strings

````bash
 % java TranslatePPT MyDeck.pptx
````

This produces a file named `MyDeck.text,`
which looks like this:

````
````

### Step 2: Edit text strings

Now you use your favorite text editor to edit
any of the text strings you want to modify, deleting
the ones you don't. (Or you can just leave them there
to be re-written as-is to the output file, although this
will slow things down for huge files.)

After I've finished making my edits, the `MyDeck.text`
file looks like this:

````bash
 
````

### Step 3: Replace text strings

Finally, you do a second run of 
<span STYLE="font-variant: small-caps;">translatepptx</span>
with the same `.pptx` file but now with
the new command-line argument `--Translations MyDeck.text`
to specify my list of revised text strings:

````bash
 % java TranslatePPT MyDeck.pptx --Translations MyDeck.text
````

This produces a new `.pptx` file called `MyDeck_Translated.pptx:`

````
````
