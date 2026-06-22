# Gesture Presentation Controller

Control PowerPoint slideshows using hand gestures via webcam. Navigate slides with swipe gestures, use your index finger as a laser pointer - no clicker needed.

## How it works

| Gesture | Action |
|---------|--------|
| Index + middle finger, swipe left | Next slide |
| Index + middle finger, swipe right | Previous slide |
| Index finger only | Laser pointer (moves on slide) |
| Fist (all fingers down) | Reset - ready for next gesture |

Uses a state machine to prevent accidental triggers. Once a gesture is recognized, it locks until you reset with a fist.

## Demo

The camera window shows:
- Hand skeleton with landmarks
- Current mode (SWIPE / POINTER / LOCKED / RESET / IDLE)
- Slide number (live from PowerPoint)
- FPS counter

## How to use

1. Open your PowerPoint and start slideshow (F5)
2. Run the script
3. Show index + middle finger → enter swipe mode
4. Swipe hand left → next slide
5. Make fist → reset
6. Show index finger only → laser pointer activates
7. Make fist → reset, ready for next gesture

## Project Structure

```
├── PresentationController.py   # Main controller
├── HandTrackingModule.py       # Shared hand detection module
├── hand_landmarker.task        # MediaPipe hand detection model
├── NOTES.md                    # Development learnings
└── README.md
```

## Setup

```bash
pip install opencv-python mediapipe pyautogui pywin32
```

- Windows required (uses PowerPoint COM API via pywin32)
- Python 3.9-3.12
- PowerPoint must be installed

## Run

```bash
python PresentationController.py
```

Press `q` to quit.

## Tech

- MediaPipe Tasks API - real-time hand landmark detection (21 points)
- OpenCV - camera feed + visual overlay
- pyautogui - sends keyboard shortcuts for slide navigation + moves cursor for pointer
- pywin32 - reads current slide number from PowerPoint via COM API (runs in background thread)
- State machine pattern - prevents gesture flickering and accidental triggers
