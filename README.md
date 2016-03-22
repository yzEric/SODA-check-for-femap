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

- - - -

## Benchmark results

4 implementations of the same mesh distortion analysis were used to compare the performances of each solutions.

The table below shows the results obtained for the same model on the same PC.

| Version              | code length* | Duration  | Speed factor  
| -------------------- | -----------: | --------: | ------------- 
| WinWrap full Femap   |   70 lines   |    300s   |  x1
| WinWrap max optim.   |  125 lines   |     20s   |  x15
| VB single thread     |  110 lines   |     11s   |  x27
| VB multi threads     |  140 lines   |      5s   |  x75 _4-CPUs_

_*Code length only include code for mesh analysis, not selection of element and results storage_

The results show that it is possible to improve performance by reducing and optimizing calls to Femap functions.

The versions compiled in VB provides a performance increase but require few modifications of the types of variables, the gain is extremely important with multi threads implementation.


- - - -

## Element quality check: SODA criteria

SODA distortion criteria is based on the sum of deviation angles. 

- For quadrilateral faces, the deviation is based on a 90 degree angle.
- For triangular faces, the deviation is based on a 60 degree angle.



| Shape     | Formula
| :-------: | ---------------------------------------------------------------- |
| ![Quad]   |   Soda_Quad = abs(90-α1) + abs(90-α2) + abs(90-α3) + abs(90-α4)
| ![Tria]   |   Soda_Tria = abs(60-α1) + abs(60-α2) + abs(60-α3)


[Quad]: https://raw.githubusercontent.com/yzEric/SODA/master/Quad.png "Soda Tria"
[Tria]: https://raw.githubusercontent.com/yzEric/SODA/master/Tria.png "Soda Tria"



## Implementations

### - WinWrap: full Femap -
This script use only integrated funtions of Femap, it is short and easy to write but it is very slow.


This script produce a lot of calls of Femap API functions

```
  For each element 
      Get element data-> elem.get(ID)
      For each corner
         mesure corner angle -> model.feMeasureAngleBetweenNodes
      Next
  Next 
  
  Number of calls:
     - Femap API = number of elements * 4 
```


### - Win Wrap: optimized -
This script use few tricks to speed up the previous version.

Few changes in code permit significantly reduce the number of calls of Femap API.


```
  Get all nodes on elements      -> nodeSet.AddSetRule( elemSet.ID , FGD_NODE_ONELEM )
  Get coordinates of these nodes -> Node.GetCoordArray
  Get data of all elements       -> Elem.GetAllArray
  
  For each element 
      For each corner
         mesure corner angle     -> non-femap function
      Next
  Next 
  
  Number of calls:
     - Femap API = 3
     - User function = number of elements * 4  
```


### - Visual Basic: single thread -
Code is pretty the same as the previous version.

Minor changes have been made to adapt the code in vb, mainly the adaptation of the type of variables.

```
  Get all nodes on elements      -> nodeSet.AddSetRule( elemSet.ID , FGD_NODE_ONELEM )
  Get coordinates of these nodes -> Node.GetCoordArray
  Get data of all elements       -> Elem.GetAllArray
  
  For each element 
      For each corner
         mesure corner angle     -> non-femap function
      Next
  Next 
  
  Number of calls:
     - Femap API = 3
     - User function = number of elements * 4   
```


### - Visual Basic: multi threads -
Code is pretty the same as the single thread version.

After obtained the data from the model, multiple threads are used to analyze data.


```
  Get all nodes on elements      -> nodeSet.AddSetRule( elemSet.ID , FGD_NODE_ONELEM )
  Get coordinates of these nodes -> Node.GetCoordArray
  Get data of all elements       -> Elem.GetAllArray
  
  # Dispatch on threads
    For each element 
        For each corner
           mesure corner angle     -> non-femap function
        Next
    Next 
  #
  
  Number of calls:
     - Femap API = 3
     - User function = number of elements * 4  
```
