String cmd = "gs -q -dNOPAUSE -dBATCH -dSAFER -sDEVICE=tiffgray -r300 -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -sOutputFile=" + tifFullName + " " + pdfFullName;
