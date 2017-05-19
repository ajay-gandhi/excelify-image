# Excelify

> Convert images to Microsoft Excel documents

## Wat

Remember those chain emails from 2005 where someone attached an Excel document?
If you zoomed out enough, the document looked like an actual image. Well, this
Node.js module does that for you.

## How

It's a lot of unnecessarily difficult XML parsing (since that's what Excel
documents really are). The target image gets iterated pixel by pixel, and each
pixel gets its own cell in the document. Resize the columns, and you're done!

