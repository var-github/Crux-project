# Task
![image](https://github.com/user-attachments/assets/f8ce2836-bf85-4797-aaf3-d8f209d5e7e5)

Note: Windows users can use Ctrl+Shift+C on a file to directly copy its path, this can be pasted anywhere in the program
#### Page Range
Whenever you want to select all pages from 20 to 30, no need to individually mention all page number, just input 20-30
Both page number 20 and 30 are included

#### Pypdf
The pypdf module is used for most of the functions like merge, split, add bookmark, scale pages, etc

#### Spire.pdf
Since pypdf cannot edit or delete pdf's, spire.pdf is used for these 2 cases
ALso due better compression algorithm that pypdf, even the compress pdf function uses spire.pdf
But the free version of Spire only allow 10 pages to be processed, so these 3 functions are limited in this way

#### Docx2pdf and pptxtopdf
These modules are used to convert docx and pptx file to pdf files

All functions except merge will save the pdf in the original directory itself, merge function will ask user which directory to store pdf in.
