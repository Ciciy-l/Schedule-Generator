# Schedule-Generator V1.6
    Generate shift schedules based on rules.

## Introduce

- Set up the list of duty personnel.
- Optional Whether to automatically skip holidays.
- Customize the prefix.
- There are two versions: command line and GUI.
- ……

## How to operate

### Command line operation
  - When entering a month, enter a pure digit from 1 to 12. Do not leave any space.
  - Excel files in XLSX format are supported only. Excel files in XLS format are not supported. Excel files, generation. exe, and personal_information must be in the same directory.
  - The file name can be copied and pasted by right-clicking, or click the icon in the upper right corner of the command line window to enter the Edit option and click Paste.
  - In personal.txt in the personal_information folder, each person's name must have an exclusive line. The number in the first line indicates the first person's name to start the shift.
  - After running, the personnel list of the next month will be automatically rearranged. At the same time, the order of the current month will be backed up in the personal.bak file, which can be opened in Notepad to view.
### GUI operation
  - Select the month to build.
  - Choose whether to skip holidays.
  - Enter the name of the template file.
  - Click on the button

## The next version

- Parameterized control of write cells.
- Add a configuration file.
- Optimized the template configuration mode.