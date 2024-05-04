# EMG Reference Generator

## Introduction
This program generates a nerve conduction study (NCS) report with reference values included and abnormal values highlighted. The input is a Microsoft Word (.docx) document which includes three tables. In order, the tables represent:
1. Motor Nerve Conduction Studies
2. F-Waves
3. Sensory Nerve Conduction Studies

All content aside from these tables is ignored by the program, allowing for other changes in formatting. An example of the expected format of the data tables is provided in `Test/sample.docx`. Any order and assortment of nerve tests within each table is acceptable, however the table should follow the specified organization. The output of the program is a Microsoft Word (.docx) document, which will be stored in the same directory as the source file with the title `emgref.docx`. The file will include three tables corresponding to the above, with reference values provided where appropriate and abnormal values highlighted in red. 

## Installation
The latest release may be downloaded [here](https://github.com/jonzia/EMG-Reference-Generator/releases). Please note that the current version supports Mac OS only. To install, download the `EMGRG.dmg` file and save the application.

## How to Use
1. Specify the source file by pressing "Select File".
2. Enter the patient's age.
3. Enter the patient's height to the nearest inch.
4. Run the program by selectingn "Generate".
5. When complete, the generated report will be available as `emgref.docx` in the same directory as the source file.
6. Select "Exit" to close the program, or repeat the above steps to generate a new report.

## Disclaimers
- This software is not intended for use in the diagnosis or management of medical conditions.
- Reference values for patients under four years of age are not included.
- The reference values in this software may be found in `Test/reference.png`
