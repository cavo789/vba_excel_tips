# VBA Tips for Excel

Some tips and tricks when writing VBA code

![Banner](./banner.svg)

* [Improved speed](#improved-speed)

## Improved speed

Make sure to set the calculation mode to manual and to disable screen updating and events.

Always first save the current state in working variables.

```vbnet
Dim xlMode As XlCalculation
Dim bEvents As Boolean, bScreenUpdating As Boolean

    xlMode = Application.Calculation
    Application.Calculation = xlCalculationManual

    bEvents = Application.EnableEvents
    Application.EnableEvents = False

    bScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = bScreenUpdating

    ' ... Put your code here ...

    Application.ScreenUpdating = bScreenUpdating
    Application.EnableEvents = bEvents
    Application.Calculation = xlMode
```
