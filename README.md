# Copy Pasta

Copy Pasta is a dockable ScriptUI panel for Adobe After Effects that lets you:

- copy selected visual layers from After Effects to the system clipboard
- paste image data from the system clipboard into the active composition

## Features

- Large `Copy` and `Paste` buttons in a clean ScriptUI panel
- Copy support for:
  - Shape Layers
  - Still Image Layers (footage)
- Paste support for images copied from browsers, design tools, and other apps
- Automatic import + layer creation when pasting into the active comp
- Clear validation and error messages for common failure states

## Requirements

- Adobe After Effects (modern versions recommended)
- OS: Windows or macOS
- After Effects preference enabled:
  - `Preferences > Scripting & Expressions > Allow Scripts to Write Files and Access Network`

## Installation

1. Copy `Copy Pasta.jsx` into your After Effects ScriptUI Panels folder.

   - Windows:
     `C:\Program Files\Adobe\Adobe After Effects <version>\Support Files\Scripts\ScriptUI Panels\`
   - macOS:
     `/Applications/Adobe After Effects <version>/Scripts/ScriptUI Panels/`

2. Restart After Effects.
3. Open the panel from `Window > Copy Pasta`.
4. Dock the panel if desired.

## Usage

### Copy from After Effects

1. Open a composition.
2. Select one supported layer (shape or still image layer).
3. Move the playhead so the selected layer is visible.
4. Click `Copy`.
5. Paste in another app (Photoshop, Figma, browser chat, etc.).

### Paste into After Effects

1. Copy an image from another app.
2. Return to After Effects and open the target composition.
3. Click `Paste`.
4. The image is imported into the project and added as a new layer in the active comp.

## Error Handling

The panel reports clear messages for:

- no active composition
- no selected layer for `Copy`
- unsupported selected layer type
- selected layer not visible at current time
- clipboard does not contain readable image data

## Notes

- Clipboard image operations are implemented with OS-level commands.
- Temporary files may be created during copy/paste operations.
- Linux is not supported for clipboard image transfer in this script.

## Troubleshooting

- If `Paste` says it cannot read an image:
  - confirm the image can be pasted in another application
  - re-copy the image and try again
  - keep the source app open while testing
- If `Copy` fails:
  - ensure the selected layer is visible at the current playhead time
  - verify the selected layer is a shape layer or still image layer
  - make sure scripting/network access is enabled in preferences

## Files

- Panel script: `Copy Pasta.jsx`
- Project documentation: `README.md`
