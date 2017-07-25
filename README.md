# MIsearch
**MIsearch** facilitates the automated searching of extracted monoisotopic masses from mass spectrometry data against a theoretical list of targets, across a defined mass bin (usually 5 ppm). For experimental data that falls within the expected mass range, matches correctly to a second column (charge state) and is identified in two channels (native and stable isotope labeled), the data header (molecule name, theoretical mass, charge) is reported. Thus, in < 1 minute, it matches the hundreds of thousands of experimental lists to a theoretical database, greatly reducing analysis time and removing the bias of checking each value manually. In doing so, it can identify novel molecules in new species that otherwise are not targeted.

**MIsearch** is used to analyze data in the following publication: [N-linked glycosite profiling and use of Skyline as a platform for characterization and relative quantification of glycans in differentiating xylem of Populus trichocarpa](https://link.springer.com/article/10.1007%2Fs00216-016-9776-5)

## Program
**MIsearch** is a program (written in Python) that reads an input Excel document, analyzes its data, and outputs a new Excel document using the following process:
1. Reads Excel document, storing the input data into 3 separate lists (experimental values, NAT intervals, and SIL intervals)
2. The NAT intervals are iterated through to see if any of the experimental values have the same charge and fall within the low/high interval range. If so, the NAT theoretical, charge, and ID are written to an output list.
3. The above process is repeated for the SIL intervals.
4. The NAT and SIL output lists are compared, looking for intervals that have the same charge and ID. These matches are written into a final output file.
5. The results are output to a second Excel document.

## Getting Started
These instructions will get you a copy of the project up and running on your local machine for analysis. **Note: These instructions assume a Windows operating system is being used.**

### Prerequisites
**MIsearch** was developed using the following programs and versions. There is no guarantee of program functionality using different versions of these programs.
* **Python** (version 2.7.6): https://www.python.org/download/releases/2.7.6/
* **setuptools** (version 18.5): https://pypi.python.org/pypi/setuptools/18.5
* **openpyxl** (version 2.3.0):  https://pypi.python.org/pypi/openpyxl/2.3.0

### Installation
1. Download and install Python.
2. Set Python Path and Path variables:
    * Go to Control Panel > System > Advanced System Settings > Environmental Variables.
    * Add user variable ```PATH``` and value ```C:\python27```.
    * Add system variable ```PYTHONPATH``` and value ```C:\python27```.
3. Download setuptools.
    * Unzip and then go into directory.
    * Run: ```python setup.py install```
4. Download openpyxl.
    * Unzip and then go into directory.
    * Run: ```python setup.py install```

## Running the Program
1. From the command line, run the program by typing: ```python MIsearch.py [file.xlsx]```
2. Multiple files may be run simultaneously by typing: ```python MIsearch.py [file.xlsx] [file.xlsx] [file.xlsx] …```
3. The program will output a new file with name ```file_results.xlsx```. It will output the ID, theoretical charge for the NAT and SIL species, and their associated charge states.

### Input File Format
| Experimental | Charge | NAT Theoretical | NAT Charge | NAT Low | NAT High | ID | SIL Theoretical | SIL Charge | SIL Low | SIL High |
|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|:---:|
|  -  |  -  |  -  |  -  |  -  |  -  |  -  |  -  |  -  |  -  |  -  |

### Output File Format
| Output-Charge | Output-NAT Theoretical | Output-SIL Theoretical | Output-ID |
|:---:|:---:|:---:|:---:|
|  -  |  -  |  -  |  -  |

## Authors
* **Jon Ziefle** - *Programming* - [jonziefle](https://github.com/jonziefle)
* **Elizabeth Hecht** - *Inspiration and scientific descriptions*

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details
