This Visual Basic Source Code Project is an implementaation of Splay Tree compression in Visual Basic.
This source code is a translation of pascal routines from splay.pas to Visual Basic. The idea is taken from
splay.pas. Please see splay.pas for details. I have included splay.pas for reference.

Any modifications of the Visual Basic Source code please give me a copy and email to fred72ph@yahoo.com.
This is my first submission here in planet source code. I hope it would help programmers interested in compression.
This routines is faster than the huffman compression implemented in Visual Basic. Compression is fair. But I
plan to modify it to give better compression ratio. I hope you enjoy it. Please vote for this code.

I implement as module so that you can include it in your programs.
To compress:
  SplayCompress(FileName);
To expand:
  SplayExpand(FileName);

Please Note: Just modify the path string in the source code to handle the compression and expansion of file.

If ever you include it in your program don't hesitate to mention my name as an acknowledgement. Thanks.
My name is found in the module.


December 23, 2002

I added a Run Length Encoding routine to preprocess the file which yields good compression ratio for files
with repeated byte sequences. I implemented the module in project for compressing and expanding single
files. Later on this will become a very big project that will include compression of multiple files. As of now
we just compress single files. All routines in the module are optimized for speed and better performance.
There are many things you can learn from this code. The implementation of Do-Loop, Do While-Loop and 
For Next Loop and some file functions. Hope you like it.  Please Vote for this code.
  e.
