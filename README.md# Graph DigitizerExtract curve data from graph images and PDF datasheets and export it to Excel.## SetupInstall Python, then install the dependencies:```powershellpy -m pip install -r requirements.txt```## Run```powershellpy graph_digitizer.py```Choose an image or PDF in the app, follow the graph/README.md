# Graph Digitizer

Extract curve data from graph images and PDF datasheets and export it to Excel.

## Setup

Install Python, then install the dependencies:

```powershell
py -m pip install -r requirements.txt
```

## Run

```powershell
py graph_digitizer.py
```

Choose an image or PDF in the app, follow the graph/curve prompts, and export the resulting workbook. The generated workbook contains a combined sheet and individual sheets for each curve.

On systems without the `py` launcher, use `python` or `python3` in place of `py`. Tkinter must also be available in your Python installation.
