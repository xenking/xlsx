# XLSX

[![Build Status](https://circleci.com/gh/xenking/xlsx.svg?style=svg)](https://circleci.com/gh/xenking/xlsx)
[![Codecov](https://codecov.io/gh/xenking/xlsx/branch/master/graph/badge.svg?token=MsQrvi4J4e)](https://codecov.io/gh/xenking/xlsx)
[![Code Report](https://goreportcard.com/badge/github.com/xenking/redis)](https://goreportcard.com/report/github.com/xenking/redis)
[![Go Reference](https://pkg.go.dev/badge/github.com/xenking/xlsx/v3.svg)](https://pkg.go.dev/github.com/xenking/xlsx/v3)
[![License](https://github.com/xenking/xlsx#license)](https://img.shields.io/badge/license-bsd-orange.svg)

## Introduction
xlsx is a library to simplify reading and writing the XML format used
by a recent version of Microsoft Excel in Go programs.

## Tutorial

If you'd like an introduction to this project try the [new tutorial](https://github.com/xenking/xlsx/blob/master/tutorial/tutorial.adoc).

## Different versions of this project

### Prior to v1.0.0

You don't want these versions ;-)

It's hard to remember exactly, but work on this library started within
a month of the first public announcement of Go, now more than a decade
ago.  It was essentially a quick hack to get data out of XLSX files at
my workplace.  Nobody but me relied on it, so it was fine to use this
brand-new language for this task. Somewhat later I decided to share
the code, and I know it was well established as an open-source project
by the time I left that job in late 2011.

Although I did do some "release" tags, versioning in Go in the early
days relied on tagging your code with the name of the Go release
(i.e. go1.2) and then `go get` would fetch that tag, if it existed,
and if not, it'd grab the master branch.

### Version 1.x.x

Version 1.0.0 was tagged in 2017 to support vendoring tools.

As of October 8th, 2019, I've branched off v1.x.x maintenance work
from master.  The master branch now tracks v2.x.x.

If you have existing code, can live with the issues in the 1.x.x
codebase, and don't want to update your code to use a later version,
then you can stick to these releases.  I mostly won't be touching this
code, but if something really important comes up, let me know.

### Version 2.x.x

Version 2.0.0 introduced breaking changes in the API.

The scope of these changes included the way `Col` elements and
`DataValidation` works, as these aspects have been built around
incorrect models of the underlying XLSX format.

See the [milestone](https://github.com/tealeg/xlsx/milestone/5) for details.

Version 2.0.1 was tagged purely because 2.0.0 wasn't handled correctly
with regard to how go modules work. It isn't possible to use 2.0.0
from a Go Modules based project.

### Version 3.x.x 
Version 3.0.0 introduces some more breaking changes in the API.  All
methods that can return an `xlsx.File` struct now accept zero, one or
many `xlsx.FileOption` functions as their final arguments.  These can
be used to modify the behaviour of the resultant struct - in
particular they replace the `...WithRowLimit` variants of those
methods with the result of calling `xlsx.RowLimit` and they add the
ability to define a custom backing store for the spreadsheet data to
be held in whilst processing.

StreamFileBuilder has been dropped from this version of the library as it has become difficult to maintain. 

## Full API docs
The full API docs can be viewed using go's built in documentation
tool, or online at [pkg.go.dev](https://pkg.go.dev/github.com/xenking/xlsx/v3).

## Contributing

We're extremely happy to review pull requests.  Please be patient, maintaining XLSX doesn't pay anyone's salary (to my knowledge).

If you'd like to propose a change please ensure the following:

- All existing tests are passing.
- There are tests in the test suite that cover the changes you're making.
- You have added documentation strings (in English) to (at least) the public functions you've added or modified.
- Your use of, or creation of, XML is compliant with [part 1 of the 4th edition of the ECMA-376 Standard for Office Open XML](http://www.ecma-international.org/publications/standards/Ecma-376.htm).

Eat a peach - Geoff
