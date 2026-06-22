# Learnings - Gesture Presentation Controller

## What I built
A gesture-based PowerPoint controller. Swipe to change slides, point to move laser pointer. No physical clicker needed.

---

## The main challenge: flickering

First version had a simple if/else checking finger states every frame. Problem: finger detection isn't perfectly stable. It flickers between states for a few frames randomly. This caused:
- Accidental slide changes
- Pointer activating/deactivating rapidly
- Swipe triggering when bringing hand back to original position

## Solution: state machine

Instead of reacting to every frame independently, I implemented states:

```
IDLE → (2 fingers) → SWIPE → (swipe detected) → LOCKED → (fist) → RESET → ...
IDLE → (1 finger) → POINTER → (fist) → RESET → ...
```

Rules:
- Once you enter a mode, you STAY until fist resets
- After a swipe triggers, it locks immediately (no reverse swipe on return)
- Fist is the universal reset - clears everything

This one design decision fixed all the flickering problems.

---

## Swipe detection approach

Track wrist x-position when entering swipe mode. Compare current position against that start position. If difference exceeds threshold → trigger.

```python
startWristX = lmList[0][1]  # recorded when entering swipe mode
diff = currentWristX - startWristX
if diff < -threshold:  # moved left enough
    # next slide
```

Threshold of 40px worked well at arm's length from camera.

Why wrist and not fingertip? Wrist is more stable - fingertips wobble but wrist position is steady during a deliberate swipe.

---

## FPS issue with PowerPoint COM API

First tried using `win32com.client` to both control slides and read slide numbers. Every COM call blocked the main loop for ~50-100ms, causing visible FPS drops.

Fix: split the responsibilities:
- **Slide changes** → `pyautogui.press('right'/'left')` - instant keyboard shortcut, never blocks
- **Slide number reading** → background thread with its own COM connection, polls every 2 seconds

```python
# Background thread - never blocks camera feed
def update_slide_info():
    pythoncom.CoInitialize()  # COM needs init per thread
    pptApp = win32com.client.Dispatch('PowerPoint.Application')
    while True:
        currentSlide = ssw.View.Slide.SlideIndex
        time.sleep(2)
```

Key learning: COM objects can't be shared across threads. Each thread needs its own `CoInitialize()` and its own `Dispatch()` call.

---

## Pointer / laser

- `pyautogui.hotkey('ctrl', 'l')` activates PowerPoint's built-in laser pointer
- Then `pyautogui.moveTo(x, y)` moves the cursor which moves the laser
- Finger camera coordinates mapped to screen coordinates: `mapX = int(ix * screenW / wCam)`

---

## Image flipping

`cv2.flip(img, 1)` - mirrors horizontally. Without this, swiping left on your hand appears as right on screen. Confusing. With flip, everything is natural like looking in a mirror.

---

## Reusing the hand tracking module

Same `HandTrackingModule.py` from the volume controller project. No changes needed. Just imported and used:
- `findHands()` - detection + drawing
- `findPosition()` - landmark coordinates
- `fingersUp()` - which fingers are up

This validated the module design - genuinely reusable across different gesture projects without modification.

---

## Things that didn't work

- Using COM API for slide changes directly → too slow, freezes video
- Frame-by-frame finger checking without state machine → too flickery
- High swipe threshold (80px) → needed too much arm movement, uncomfortable
- Tracking fingertip for swipe direction → too jittery, wrist is more stable


