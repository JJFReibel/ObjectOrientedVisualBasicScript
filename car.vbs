' MIT License

' Copyright (c) 2020 Jean-Jacques François Reibel

' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:

' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.

' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

' This sample requires VBS 5 or higher
Class Car

   Dim wheels
   Dim doors
   Dim cylinders

   Public Default Function Init(wheelsIn, doorsIn, cylindersIn)
         wheels = wheelsIn
         doors = doorsIn
         cylinders = cylindersIn
         Set Init = Me
   End Function

   Public Property Let addWheels(wheelsIn)
      wheels = wheels + wheelsIn
   End Property

   Public Property Let addDoors(doorsIn)
      doors = doors + doorsIn
   End Property

   Public Property Let addCylinders(cylindersIn)
      cylinders = cylinders + cylindersIn
   End Property

   Public Property Let deleteWheels(wheelsIn)
      wheels = wheels - wheelsIn
   End Property

   Public Property Let deleteDoors(doorsIn)
      doors = doors - doorsIn
   End Property

   Public Property Let deleteCylinders(cylindersIn)
      cylinders = cylinders - cylindersIn
   End Property

   Public Sub printInfo
      wscript.echo "Wheel check: " & cstr(wheels)
      wscript.echo "Door check: " & cstr(doors)
      wscript.echo "Cylinder check: " & cstr(cylinders) & vbCrLf
   End Sub

End Class

wscript.echo "Creating car." & vbCrLf
Dim subaru : Set subaru = (New Car)(4,4,4)

wscript.echo "Adding wheel directly to car object."
subaru.wheels = 5
subaru.printInfo()
wscript.echo "Removing wheel using object method."
subaru.deleteWheels(1)
subaru.printInfo()
