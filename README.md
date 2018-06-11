## How to run

Install pyodbc version 4.0.22 or greater

For a missed collection notification, run:

```console
py -3 gen_html.py
```

For a change in rounds notification, run:

```console
py -3 new_rounds_gen_html.py
```

## How it works

Generates an HTML file and uses [wkhtmltopdf](https://wkhtmltopdf.org/) to convert it to a PDF