# spv2apa

### Function
This script converts SPSS outputs to APA style tables. It is very similar in purpose to the "outreg" command in Stata.

### Supported outputs
Currently supported outputs are:
- Pearson & Spearman correlations (correlation matrix)
- GLM (parameter estimats)
- Linear Regression (coefficients)

Two different thresholds for significance can be used:
- 0.01, 0.05, 0.1,
- 0.001, 0.01, 0.05 (this is more common but field specific)

### How to use
Copy one of the supported tables from the SPSS output, as plain text, to an Excel file named "input.xlsx". Paste it in A1. Run the script.
A batch file is provided for convenience. The output files will be stored as "output.docx".

The "template.docx" file can be editted to change table styling. This is done by editing the "Table Grid" style in the document. Always leave this document empty, otherwise its contents will also be copied to the output.

Want to change the significance thresholds? In the script file, look for the "significance" variable and change the assigned value between 1 and 2.
