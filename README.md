EdoliAddin

# Table of Contents
1. [Shape](#shape)
1. [Align](#align)
1. [Curve](#curve)
1. [Animation](#animation)
1. [Shortcuts](#shortcuts)

# Shape

Toggle line arrow

![icon_begin_arrow_toggle](EdoliAddin/Resources/icon_begin_arrow_toggle.png)
![icon_end_arrow_toggle](EdoliAddin/Resources/icon_end_arrow_toggle.png)

Size up arrow

![icon_begin_arrow_size_up](EdoliAddin/Resources/icon_begin_arrow_size_up.png)
![icon_end_arrow_size_up](EdoliAddin/Resources/icon_end_arrow_size_up.png)

Size down arrow

![icon_begin_arrow_size_down](EdoliAddin/Resources/icon_begin_arrow_size_down.png)
![icon_end_arrow_size_down](EdoliAddin/Resources/icon_end_arrow_size_down.png)

Connect shapes by lines

![icon_connect_shape_by_lines](EdoliAddin/Resources/icon_connect_shape_by_lines.png)

Invert image

![icon_image_invert](EdoliAddin/Resources/icon_image_invert.png)

Time image

![icon_image_trim](EdoliAddin/Resources/icon_image_trim.png)


# Align

Place labels on the bottom/left/right/top side of images

![icon_label_bottom](EdoliAddin/Resources/icon_label_bottom.png)
![icon_label_left](EdoliAddin/Resources/icon_label_left.png)
![icon_label_right](EdoliAddin/Resources/icon_label_right.png)
![icon_label_top](EdoliAddin/Resources/icon_label_top.png)

Transpose shapes

![icon_transpose](EdoliAddin/Resources/icon_transpose.png)

Group images and labels

![icon_label_group](EdoliAddin/Resources/icon_label_group.png)

Align shapes with previous slide

![icon_align_prev_slide](EdoliAddin/Resources/icon_align_prev_slide.png)

Align shapes with next slide

![icon_align_next_slide](EdoliAddin/Resources/icon_align_next_slide.png)

Swap multiple shapes

![icon_swap_cycle](EdoliAddin/Resources/icon_swap_cycle.png)
![icon_swap_cycle_reverse](EdoliAddin/Resources/icon_swap_cycle_reverse.png)

Align shapes in diagonal

![icon_snap_diag_downright](EdoliAddin/Resources/icon_snap_diag_downright.png)
![icon_snap_diag_upright](EdoliAddin/Resources/icon_snap_diag_upright.png)

Align shapes in grid automatically

![icon_align_grid](EdoliAddin/Resources/icon_align_grid.png)

Align shapes in grid with custom padding and column. Shapes are placed in row major order. The shapes are sorted in the selected order.

![icon_grid](EdoliAddin/Resources/icon_grid.png)


# Curve

Create polyline of equation

![icon_polyline_of_equation](EdoliAddin/Resources/icon_polyline_of_equation.png)

Create bezier curve of equation

![icon_curve_of_equation](EdoliAddin/Resources/icon_curve_of_equation.png)

This addin uses [expressive](https://github.com/bijington/expressive) for parsing equation. Use parameter `t` for drawing curve. Also the range should be set for `t`. There are some examples


```
Range: 0 - 10*PI
X: t
Y: Cos(t)
```
![t_cos](images/t_cos.png)

```
Range: 0 - 2*PI
X: Cos(t)
Y: Sin(t)
```
![cos_sin](images/cos_sin.png)

```
Range: 0 - 2*PI
X: Cos(3*t)
Y: Sin(2*t)
```
![3cos_2sin](images/3cos_2sin.png)


# Shortcuts

## Align
<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD2</kbd>: Align bottom

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD4</kbd>: Align left

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD6</kbd>: Align right

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD8</kbd>: Align top

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD5</kbd>: Align center

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>H</kbd>: Align center horizontal

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>T</kbd>: Align center vertical

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD7</kbd>: Align in row

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>NUMPAD1</kbd>: Align labels to bottom

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>HOME</kbd>: Bring to front

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>END</kbd>: Send to back

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>PAGEUP</kbd>: Bring forward

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>PAGEDOWN</kbd>: Send backward

<kbd>CTRL</kbd>+<kbd>NUMPAD2</kbd>: Align bottom of

<kbd>CTRL</kbd>+<kbd>NUMPAD4</kbd>: Align left of

<kbd>CTRL</kbd>+<kbd>NUMPAD6</kbd>: Align right of

<kbd>CTRL</kbd>+<kbd>NUMPAD8</kbd>: Align top of

## Shape

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>0</kbd>: Toggle line

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>+</kbd>: Thickening line width

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>-</kbd>: Thinning line width

<kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>;</kbd>, <kbd>CTRL</kbd>+<kbd>ALT</kbd>+<kbd>'</kbd>: Change line dash style

## Object naming for animation

<kbd>CTRL</kbd>+<kbd>SHIFT</kbd>+<kbd>1</kbd>: Set object name to `!!ID1`

<kbd>CTRL</kbd>+<kbd>SHIFT</kbd>+<kbd>2</kbd>: Set object name to `!!ID2`

<kbd>CTRL</kbd>+<kbd>SHIFT</kbd>+<kbd>3</kbd>: Set object name to `!!ID3`

...
