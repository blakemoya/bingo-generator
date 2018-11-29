# Bingo-generator
Bingo-generator is an educational resource for creating quizzes formatted as bingo games. Given a list of questions in a table in a docx file, Bingo-generator can build a specified number of unique cards in an output docx file. The intention is that studnts will have to solve the problems on their bingo card before the answers are announced to determine if they have won.
## Getting Started
These instructions will show you how to generate your own bingo cards
### Prerequisites
- numpy
- python-docx
- PyQt5
- pyinstaller (optional)
### Installation
Clone this repo
```
git clone https://github.com/blakemoya/bingo-generator
```
Optionally the script can be exported to an executable calling
```
pyinstaller bingo-generator.py --onefile --noconsole
```
**However** the executable will fail if it is not run with an empty file named "template.docx" in the same directory.
### Usage
1. Run
```
python bingo-generator.py
```
(Or run the executable if you chose to export)

2. A UI will appear that prompts for 
    - input file
    - number of cards to be generated
    - output directory
3. **The current version does not validate this input**, so be careful.
  - The input file must be a docx file containing a table with at least 24 rows whose first columns contains the questions to be included in your bingo cards (as in input.docx)
  - The number of cards to be generate must be less than the number of cards that can possible be created (which is the number of combinations of 24 of the number of rows in the table).

4. Once you have specificed your input, click Generate.

5. The Generate button will change to say Done, you may now exit the program.

**Note:** If you press Done, the program **will** execute again, with whatever input is currently showing to be selected.

