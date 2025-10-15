# Lovable Progress Bar for PowerPoint

A VBA macro to add a customizable progress bar to your PowerPoint slides. It can display as a **single continuous bar** or a **multi-step bar**, optionally with the last step having a different color.

## Features
- Single or multi-step progress bar
- Customizable height, color, transparency, corner radius, and margin
- Multi-step mode allows configurable gap and different last box color
- Automatically removes previous progress bars before drawing
- Supports offsets for start and end slides

## How to Use

### 1. Enable Developer Tab
1. Open PowerPoint
2. Go to `File` → `Options` → `Customize Ribbon`
3. Check `Developer` in the right-hand list
4. Click `OK`

### 2. Open VBA Editor
1. Click `Developer` → `Visual Basic` or press `ALT + F11`
2. In VBA editor, go to `Insert` → `Module`
3. Paste the `Lovable_Progress_Bar` macro code

### 3. Run Macro
1. Close VBA editor
2. Go to `Developer` → `Macros`
3. Select `Lovable_Progress_Bar` → Click `Run`

### 4. Customization
- When prompted, you can use default values or input your own:
  - `startOffset`: number of slides to skip at start
  - `endOffset`: number of slides to skip at end
  - `barColor`: RGB color of the bar
  - `barHeight`: height in pixels
  - `transparency`: 0 (solid) to 1 (fully transparent)
  - `cornerRadius`: rounded corner radius
  - `margin`: distance from slide edges
- Choose mode: `single` or `multi`
- If `multi`:
  - `gap`: space between boxes
  - `lastBoxDifferentColor`: whether the last box has a different color
  - `lastBoxColor`: color of the last box if different

### Notes
- Previous bars (LPB_) are removed before drawing new ones.
- Single mode ignores multi-step settings.
- Works on all slides in the active presentation.

## License
MIT License