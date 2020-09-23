Attribute VB_Name = "Module1"
Option Explicit

'This one is the optimised routine
Public Sub PauseOptimised1(ByVal dblInterval As Double)
   'ByVal simplifies becuase VB know that it doesn't have to return a value
   'Type cast variable eliminates the coersion and Val command
  Dim dblEndTime As Double         'typecast Dim smaller memory footprint
  dblEndTime = Timer + dblInterval 'Do math only once (NB addition is also faster than subtraction)
  Do While Timer < dblEndTime      'Do/Loop runs faster without the unnecessary math
    DoEvents                       'so the sub responds faster/more accurately
  Loop

End Sub

'This one is a differently optimised version
Public Sub PauseOptimised2(ByVal dblInterval As Double)
'reversing the test runs just a bit faster on my system but not consistently
'and only for long pauses
  Dim dblEndTime As Double

  dblEndTime = Timer + dblInterval
  Do While dblEndTime > Timer
    DoEvents
  Loop

End Sub

'This one is the original unoptimised routine
Public Sub PauseWeak(interval)                       'default Variant has large memory footprint
  Dim Current                                        'default Variant has large memory footprint
  Current = Timer                                    'Coerce variant to Long
  Do While Timer - Current < Val(interval)           'Doing unnecessary math in loop
    DoEvents                                         'and pointless multiple conversions of numeric Variant to Numeric
  Loop                                               ' slows response time

End Sub


Public Sub PauseLong(ByVal lngInterval As Long)
'NOTE this will consistently return early
'VB coerces the Timer return value (Double) to Long in the Do line
' assuming lngEndTime = 10 then when Timer returns 9.5 (rounds down to 9) the loop continues but
' when Timer = 9.5000000000001 it rounds up to 10 and the loop exits
  Dim lngEndTime As Long

  lngEndTime = Timer + lngInterval
  Do While Timer < lngEndTime
    DoEvents
  Loop

End Sub
':)Code Fixer V4.0.27 (Friday, 20 January 2006 00:15:48) 1 + 47 = 48 Lines Thanks Ulli for inspiration and lots of code.
':)SETTINGS DUMP: 133302322223333232|33332222222222222222222222222222|1112222|2221222|222222222233|111111111111|1122222222220|333333|
