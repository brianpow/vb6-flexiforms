#Creating Flexible Forms in Visual Basic (Flexi-Forms)#

## Introduction ##
Ever designed a Form with List or TreeView controls that looked okay, but turned out to be OTT if there wasn't much data to be displayed, or felt cramped if there was plenty to be displayed? Thought about making the Form resizeable? Then the pain really starts... when dynamically resizing and repositioning the controls.

Well, here's a class that'll take the pain out of this. It can be applied to ANY form, by adding just a couple of extra lines of code, and making use of the 'Tag' control property.

Add the `clsLayout` and `clsControl` class files to your project, and then add these few lines to the code for the Form.

    Private myLayout As New clsLayout

    'Many codes...

    Private Sub Form1_Load()
    	myLayout.SetLayout Me
    End Sub

    'Many codes...

    Private Sub Form1_Resize()
    	myLayout.RedrawLayout
    End Sub

The last step is to specify how you want the controls on the Form to align¡Xsimply specify the layout in the `` `Tag` `` property of the controls. You can use a mixture of fixed and variable layout properties:

| Property  | Description                                               |
|-----------|-----------------------------------------------------------|
| None      | No alignment                                              |
| fixLeft   | Control alignment fixed towards left of Form              |
| fixTop    | Control alignment fixed towards top of Form               |
| fixRight  | Control alignment fixed towards right of Form             |
| fixBottom | Control alignment fixed towards bottom of Form            |
| varX      | Control x position adjusted proportionally to form width  |
| varY      | Control y position adjusted proportionally to form height |
| varWidth  | Control width adjusted proportionally to form width       |
| varDepth  | Control depth adjusted proportionally to form depth       |


The following are combinations of layout styles:

| Style         | Description                                                            |
|---------------|------------------------------------------------------------------------|
| varHorizontal | Left + Right (control width proportional proportional to Form width)   |
| varVertical   | Top + Bottom (control height proportional proportional to Form height) |
| varFull       | Left + Right + Top + Bottom (control size dependent on Form size)      |

Here's a useful example (see screen shots):

    Label1.Tag ""
    List1.Tag "varWidth varVertical"
    Label2.Tag "varWidth fixRight"
    List2.Tag "varWidth fixRight varVertical"
    cmdOk.Tag "fixRight fixBottom"
    cmdCancel.Tag "fixRight fixBottom"

## Additional Notes on Usage ##
---
The layout names are case sensitive (actually, the code scans for single uppercase letters). You may also embed controls within controls (for example, put a List control in a Frame). However, you should not use the Bottom or Right settings for a control within a frame that itself has the Bottom or Right setting. It doesn't make sense to use the Horizontal setting on two side-by-side controls, nor use the Vertical setting on two vertically stacked controls. The controls will overlap when the form is resized.
How It Works

The `SetLayout` function stores the minimum width and height of the form, and creates a collection of `clsControl` objects. A `clsControl` contains a reference to a control on the `Form`, together with layout information¡Xindentation from top, bottom, left, and right sides of the `Form`, and the layout style (obtained from the ``'Tag'`` property).

The `RedrawLayout` function is then invoked from the `Form_Resize()` event, which repositions and/or resizes the controls in this collection. It does this using the control's indentation information and layout style, in relation to its parent's current width and height.

## Prerequisites ##
- Visual Basic 6.0

## Authors ##
* Richard Stockdale

## Changes ##
* 2003-01-27
    * Original version by Richard Stockdale

[Flexible Forms](http://www.codeguru.com/vb/gen/vb_forms/resizing/article.php/c5949/Creating-Flexible-Forms-in-Visual-Basic-FlexiForms.htm)
