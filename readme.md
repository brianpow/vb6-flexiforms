#Creating Flexible Forms in Visual Basic (Flexi-Forms)#
A fork of [Flexible Forms](http://www.codeguru.com/vb/gen/vb_forms/resizing/article.php/c5949/Creating-Flexible-Forms-in-Visual-Basic-FlexiForms.htm) with [TypeLib](http://stackoverflow.com/questions/1847411/control-properties-in-visual-basic-6)

## Introduction ##
When working on old VB6 project, a common pain in ass is realigning and calculating the new size of controls when form resize. VB6, which is released in 1998, does not provide grid sizer or layout management as modern languages do. One day, I accidentally found [Flexible Forms] and highly impressed by its clean and direct implementation. By simply adding a magic words in tag, your controls will be aligned and resized automatically!

## Usage ##
Add the `clsLayout` and `clsControl` class files, and `mType` module to your project, and then add these few lines to the code for the Form.

    Private myLayout As New clsLayout

    'Many codes...

    Private Sub Form1_Load()
    	myLayout.SetLayout Me
    End Sub

    'Many codes...

    Private Sub Form1_Resize()
        myLayout.RedrawLayout 'Add True if you wants to resize when form is smaller than original size
    End Sub

The last step is to specify how you want the controls on the Form to align¡Xsimply specify the layout in the `` `Tag` `` property of the controls. You can use a mixture of fixed and variable layout properties:

| Property  | Abbreviation | Description                                               |
|-----------|--------------|-----------------------------------------------------------|
| None      |              | No alignment                                              |
| fixLeft   | L            | Control alignment fixed towards left of Form              |
| fixTop    | T            | Control alignment fixed towards top of Form               |
| fixRight  | R            | Control alignment fixed towards right of Form             |
| fixBottom | B            | Control alignment fixed towards bottom of Form            |
| varX      | X            | Control x position adjusted proportionally to form width  |
| varY      | Y            | Control y position adjusted proportionally to form height |
| varWidth  | W            | Control width adjusted proportionally to form width       |
| varDepth  | D            | Control depth adjusted proportionally to form depth       |

The following are combinations of layout styles:

| Style         | Abbreviation | Description                                                                   |
|---------------|--------------|-------------------------------------------------------------------------------|
| varHorizontal | H            | fixLeft + fixRight (control width proportional proportional to Form width)    |
| varVertical   | V            | fixTop + fixBottom (control height proportional proportional to Form height)  |
| varFull       | F            | fixLeft + fixRight + fixTop + fixBottom (control size dependent on Form size) |

## Prerequisites ##
* Visual Basic 6.0
    * Please add TypeLib in the Reference

## Authors ##
* Richard Stockdale
* Brian Pow

## Changes ##
* 2003-01-27
    * Original version by Richard Stockdale
    
* 2015-02-12
    * Improve and support all ScaleMode
    * Fix a bug that Controls inside container (e.g. PictureBox/Frame) are not resized and aligned correctly.
    * Add a switch to `RedrawLayout` to allow resizing if form is smaller than original size

* 2015-02-24
    * Controls in SSTab now work
    * Minor bug fix

## Known Issues ##
* The spacing between controls will also expand and shrink when form resize (e.g. varWidth, varDepth). 

## References ##
[Flexible Forms](http://www.codeguru.com/vb/gen/vb_forms/resizing/article.php/c5949/Creating-Flexible-Forms-in-Visual-Basic-FlexiForms.htm)
[TypeLib](http://stackoverflow.com/questions/1847411/control-properties-in-visual-basic-6)