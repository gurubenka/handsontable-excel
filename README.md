## Handsontable-excel

This plugins enables Excel-like formula evaluation for [Handsontable](https://github.com/warpech/jquery-handsontable).

## Demo
See demo on [JSFiddle](http://jsfiddle.net/uszywieloryba/7c5dD/6/)

## Usage
Include *handsontable-excel* code after *handsontable* code.

```html
<script src="lib/jquery.min.js"></script>
<script src="dist/jquery.handsontable.full.js"></script>
<link rel="stylesheet" media="screen" href="dist/jquery.handsontable.full.css">

<script src="jquery.handsontable-excel.js"></script>
```

## Formula Syntax

Every cell beginning with ``=`` will be evaluated. Supported syntax:

* algebraic operations
  
  ``=(4+5+6)*7``

* cell references (letter for colums, number for row format: A4, B1, C7)
  
  ``=(A1+A2+A3)*7``

* function calls

  ``=SUM(A1:A3)*7``
