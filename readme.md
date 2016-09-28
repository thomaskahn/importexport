TODO
====
# reads data in chunks of a specific size, considers a completely empty chunk as a terminal boundary. Read entire columns at a time, or groups of 2x3, etc.
->setScanBlockSize($width, $height, $verticalDirection, $horizonalDirection) 

# read the whole sheet:
->toArray()

# to provide a function which computes the effective column titles for the given dataset:
->getColumnTitles(function($record) {... return $titles; })

# to get the currently-defined column titles, use this:
->getColumnTitles() 

->onReadData(function($record) {... import the $record ... })
->parse() runs everything, blocks until completed.

->stopParsing()
->seek
->setError()
->getErrors()


