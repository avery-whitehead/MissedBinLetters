## How to run

Install pyodbc version 4.0.22 or greater

From the command line, run:

```console
py -3 gen_html.py
```

## How it works

Generates an HTML file and uses [wkhtmltopdf](https://wkhtmltopdf.org/) to convert it to a PDF

Uses the [paper.css](https://github.com/cognitom/paper-css) library for easy handling of paper sizes