Attribute VB_Name = "Work_Output_Display_Module"
  Option Explicit

' WORK OUTPUT DISPLAY MODULE
'
' MODULE VERSION: 2001.0530.0542
'
' These functions are used to print out the computations to the yellow display
' area by adding 3 simple custom commands and two objects to "Form1".
' The objects are simply a multi-line, yellow scrolling text box and a CLEAR
' button.  The yellow text box is simply given the name "Work".
'
' -----------------------------------------------
' The custom commands defined by this module are:
'
' CLEAR  - Clear out the yellow display area.
' PRT    - Print out an expression or string to the next available line.
' BLIN   - Print a blank line at the next available line position.
'
' ------------------------------------------------------------------------------
' Each time a line is printed to the display area, it uses the next available
' line after the previously printed line.
'
' Each line is simply an accumulated string of computations with the standard
' vbCrLf at the end of each line.
'
' ==============================================================================
' Clear out the yellow work display area.

  Public Sub CLEAR()
  Form1.Work.Text = ""
  End Sub

' ------------------------------------------------------------------------------
' Print some string or expression to the next available free line.
' The is no need for the () around the expression to be printed, since it is
' not a function, but a command.

  Public Sub PRT(Expression)
  Form1.Work.Text = Form1.Work.Text & Expression & vbCrLf
  End Sub

' ------------------------------------------------------------------------------
' Print a blank line.

  Public Sub BLIN()
  Form1.Work.Text = Form1.Work.Text & vbCrLf
  End Sub

' ------------------------------------------------------------------------------

