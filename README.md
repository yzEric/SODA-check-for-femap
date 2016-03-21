# SODA
_Sum Of Deviation Angles_

```
The MIT License (MIT)

Copyright (c) 2016 Eric LE GAL

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## Subject
Subject of this project is to compare performances of different implementations of a same Femap macro.
All implementations give the same results but with very differents execution speeds.

This scripts and programs do a mesh distortion analysis.


## Benchmark results

4 implementations of the same mesh distortion analysis were used to compare the performances of each solutions.

The table below shows the results obtained for the same model on the same PC.

| Version                 | code length  | Duration  | Speed factor  
| ----------------------- | -----------: | --------: | ----------- 
| WinWrap full Femap      |  70 lines    |    300s   |  x1
| WinWrap max optim.      | 125 lines    |     20s   |  x15
| VB single thread        | 110 lines    |     11s   |  x27
| VB multi threads (4CPU) | 140 lines    |      5s   |  x75


The results show that it is possible to improve performance by reducing and optimizing calls to foo functions.

The versions compiled in VB provides a performance increase but require few modifications of the types of variables, the gain becomes extremely important multi threads implementation.


- - - -
## Element quality check: SODA criteria

SODA distortion criteria is based on the sum of deviation angles. 

- For quadrilateral faces, the deviation is based on a 90 degree angle.

![Quad](https://raw.githubusercontent.com/yzEric/SODA/master/Quad.png "Soda Tria")

    Soda_Quad = abs(90-α1) + abs(90-α2) + abs(90-α3) + abs(90-α4)


- For triangular faces, the deviation is based on a 60 degree angle.

![Tria](https://raw.githubusercontent.com/yzEric/SODA/master/Tria.png "Soda Tria")

    Soda_Tria = abs(60-α1) + abs(60-α2) + abs(60-α3)

- - - -

## Implementations
### - WinWrap Full Femap -
This script is as short and as simple as possible.

The only optimisation is an intermediate function to reduce the number calls of feAppStatusUpdate


### - Win Wrap optimized -



### - Visual Basic: single thread -



### - Visual Basic: multi threads -










