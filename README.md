# clj-xlsxio

[xlsxio](https://github.com/brechtsanders/xlsxio) bidings for clojure using JNA

[![Clojars Project](https://img.shields.io/clojars/v/clj-xlsxio.svg)](https://clojars.org/clj-xlsxio)

# Why?

Java (and therefore clojure) has lots of libraries to deal with xlsx files, but all of them depends on Apache POI one way or another.
Apache POI is a HUGE library, it takes a lot of memory and it is too smart, in a bad way.
Apache POI gets in the user way a lot of times, it tries to behave like exactly like Excel, even if the user don't want to.
Parsing dates and numbers with Apache POI is MUCH HARDER than it should be, it reads the spreadsheet locale and will format everything
to behave like Excel on that region, numbers a lot of times are converted to scientific notation, it is hard to know where Apache POI will give dates on MM/dd/yyyy, dd/MM/yyyy, d/M/yy, M/d/yy and etc.
It is always a pain in the ass waste 90% of the time with this types of things instead of working on real useful code.

This library is very dumb, which sometimes that is the best.
Everything will be returned as the exact string that is written on the xlsx.
Numbers will be returned as strings and will never be converted to a different formatting and dates will be
returned in Excel timestamp (which can be easily converted to UNIX timestamp and therefore any date format you want).

This library was built using [JNA](https://github.com/java-native-access/jna) and the
excellent [xlsxio](https://github.com/brechtsanders/xlsxio) C library.
The xlsxio is being deployed with is dependencies on this library .jar file.
The xlsxio library depends on zlib, expat and minizip, if you are having problems loading these libraries when you require clj-xlsxio,
consider installing them on your system using your favorite package manager.

## Usage

This library will try to be close as possible to the data.csv interface.

```clojure
(require '[clj-xlsxio.read :refer :all])

(read-xlsx "/path/to/your.xlsx")
;=> (["row0-cell0" "row0-cell1"] 
;     ["row1-cell0" "row1-cell1"]
;     ...)
```

`read-xlsx` will return a lazy-seq of rows.

There is also the low level interface that is exactly like the xlsxio C interface.
Also check the examples folder for more info on how to use the library.

# TODO
* add the shared objects files for other OS and others architectures.

## License

Copyright © 2019 FIXME

This program and the accompanying materials are made available under the
terms of the Eclipse Public License 2.0 which is available at
http://www.eclipse.org/legal/epl-2.0.

This Source Code may also be made available under the following Secondary
Licenses when the conditions for such availability set forth in the Eclipse
Public License, v. 2.0 are satisfied: GNU General Public License as published by
the Free Software Foundation, either version 2 of the License, or (at your
option) any later version, with the GNU Classpath Exception which is available
at https://www.gnu.org/software/classpath/license.html.
