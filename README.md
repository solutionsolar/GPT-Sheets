# AI Functions for Google Sheets

This repository contains a set of AI functions for Google Sheets, which enables users to perform various operations such as summarizing text, expanding on text, analyzing data, and more using GPT-4, Davinci3, and Turbo models.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Functions](#functions)
- [Examples](#examples)
- [License](#license)

## Installation

1. Open a Google Sheets document.
2. Click on `Extensions` in the menu bar.
3. Choose `Apps Script`.
4. Copy and paste the code from this repository into the `Code.gs` file.
5. Replace `SECRET_KEY` with your OpenAI API key. Or paste your key into the yellow cell on the KeySheets tab.
![image](https://user-images.githubusercontent.com/122757410/235488835-f5d07f93-67e1-4093-a300-b6b2f06e3c4c.png)
6. Save the script by clicking on the floppy disk icon or pressing `Ctrl + S`.
7. Refresh your Google Sheets document.

## Usage

1. Open a Google Sheets document where the AI Functions code has been installed.
2. Select the cell where you want to use the AI function.
3. Type the desired AI function in the cell using the proper syntax (refer to the [Functions](#functions) section).
4. Press `Enter`.

## Functions

The following AI functions are available:

1. `Bulletize(model, text, precontext, postcontext)`
2. `Summarize(model, text, precontext, postcontext)`
3. `Detail(model, text, precontext, postcontext)`
4. `Sentiment(model, text, precontext, postcontext)`
5. `Categorize(model, text, categories)`
6. `FormulaHelper(model, text, explain)`
7. `Expand(model, text, precontext, postcontext)`
8. `Analyze(model, text, precontext, postcontext)`
9. `Direct(model, text, precontext, postcontext)`

For detailed information on each function's usage and parameters, refer to the comments in the code.

## Examples

- Summarize text using the GPT-4 model:

  ```
  =Summarize("GPT4", "This is a long text that needs to be summarized.")
  ```

- Get the sentiment of a text using the Turbo model:

  ```
  =Sentiment("Turbo", "I am really happy today.")
  ```

- Categorize text using the Davinci3 model and a list of categories:

  ```
  =Categorize("Davinci3", "This is a mystery novel.", "A1:A3")
  ```
  
  ## AI Functions Sidebar

In addition to entering your requests directly into the sheet, you can also make use of the custom sidebar to interact with the AI. To use the sidebar, simply open it by clicking on 'AI Functions' in the menu bar. Once open, fill out the form with the desired function, model, input cell, precontext, postcontext, and output cell. Choose from various precontext and postcontext options to provide the AI with additional instructions. After filling out the form, click on the 'Run Function' button. The function will then execute, and the result will be written to the specified output cell. The sidebar provides a user-friendly interface to interact with the AI, making it even easier to work with your data and analyze it in various ways.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
