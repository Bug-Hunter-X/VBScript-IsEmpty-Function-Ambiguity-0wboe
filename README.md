# VBScript IsEmpty Function Ambiguity

This example demonstrates an uncommon error in VBScript related to the `IsEmpty` function.  The `IsEmpty` function in VBScript doesn't distinguish between an empty string ("") and a variable that hasn't been assigned a value (Uninitialized). This can lead to subtle bugs if your logic depends on differentiating between these two states. 

The `bug.vbs` file showcases the problem. The `bugSolution.vbs` file provides a fix.