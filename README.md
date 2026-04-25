# Study Schedule Manager

Study Schedule Manager is a Python script for planning topic reviews over time.

The central idea is **spaced repetition**: instead of reviewing everything with the same frequency, you review a topic soon after first studying it, then gradually leave more time between reviews. The goal is to revisit material before you forget it, without wasting time repeating it too often.

This project combines that idea with a simple rating system based on previous study sessions:

- each topic starts with a predefined review schedule
- after each review, you rate how that session went from `1` to `4`
- that rating changes the spacing of the remaining reviews
- the next dates are recalculated from the date you actually reviewed

So the schedule is not fixed once at the beginning. It adapts to how well each topic seems to be sticking.

## How it works

The default schedule is based on:

```python
RANGES = [0, 7, 20]
```

This means a topic is reviewed immediately, then after 7 days, then 20 days later.

When you complete a review, the script stores both the date and a rating. That rating changes the spacing of the remaining reviews:

- lower ratings bring the next reviews closer
- higher ratings push them further away

The scheduling logic is in [data_input.py](/C:/Users/ltumi/OneDrive/CLOUD/CODE/Python/study_schedule/data_input.py).

## Requirements

- Python 3.x
- `openpyxl`

## Usage

Install the dependency:

```bash
pip install openpyxl
```

Update `FILE_PATH` in `data_input.py`, then run:

```bash
python data_input.py
```
