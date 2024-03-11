# cover-letter-generator
Generate a one page cover letter, given a word template and parameters passed in via external text file. 

Appends cover letter to existing resume with Merge-CoverLetterAndResume.ps1 script.

# TODO

- [ ] Add logic to fix `String parameter too long` when using $Document.Content.Find.Execute on line 82. Try using a different method or parsing a long string into many
