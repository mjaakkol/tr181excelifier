# tr181excelifier
Command-line tool for converting TR-181 model into Excel

# Intro
This is a little tool to that converts TR-181 XML from https://cwmp-data-models.broadband-forum.org/ into human readable Excel file. It separated Device model and profile into separate sheets. There are limited formatting and the current implementation doesn't handle markdown language used by the description yet. Only TR-181 has been tested but others should work as well.

# Usage
usage: python tr181excelifier.py [-h] -f FILE [-o OUTPUT]

If no output file is provided, the outcome is written int output.xlsx.
