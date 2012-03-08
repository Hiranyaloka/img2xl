# IMAGES2XLS SCRIPT #

IMAGES2XLS is a perl script which creates a new spreadsheet from a directory of images. The first column of the spreadsheet is a list of thumbnails. The second column holds the dirified image names, linked to a larger "thumbnail" of the same image. The spreadsheet, and smaller and larger thumbs are saved to the current working directory. If multiple directories are passed to the command line, then each (dirified) filename is appended with a a two digit number identifying the directory. Each directory produces a separate worksheet.

## DEPENDENCIES ##
- String::Dirify
- Image::Magick
- Spreadsheet::WriteExcel
- File::Find (core)
- File::Spec (core)
- Cwd (core)
## CHANGELOG ##
- version 0.1  Initial release.

## SUPPORT ##
Please send questions, comments, or criticisms to rick@hiranyaloka.com.

## COPYRIGHT AND LICENSE ##

This program is free software; you can redistribute it and/or modify it
under the terms of either: the GNU General Public License as published
by the Free Software Foundation; or the Artistic License.

See http://dev.perl.org/licenses/ for more information.

This software is offered "as is" with no warranty.

IMAGES2XLS is Copyright 2011, Rick Bychowski, rick@hiranyaloka.com.
All rights reserved.
