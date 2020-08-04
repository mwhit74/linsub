# linsub
A linear algebra VBA module

Includes:
- LU decomposition with partial pivoting (Strang)
- Forward and back substitution routines
- Matrix-vector multiplication
- Vector-matrix multiplication
- Matrix-matrix mutliplication

Motivation:
Excel/VBA does not have an LU decompisition tool built-in that I could easliy implement within VBA.
There is the MINVERSE and the MMULT worksheet functions that you *can* call from the VBA side but
typically in applied science, engineering, and mathematics we don't want or need the full inverse 
(not to mention there are lots of "wasted" operations arriving at the full inverse) we want the LU
decomposition, it is much more useful. The way you call the worksheet functions is also awkward and
not intuitive in VBA or in a programming context. 

As much as I dislike Excel/VBA and think that there are much, much better solutions out there, there's 
no getting around the fact that it is an entrenched standard within companies the world over. 

Caveats:
I no longer have easy access to a Windows system so further updates and testing might get interesting.
