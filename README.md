# Financial Fallen Angels VBA

![Apache-2.0 License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)

## Description

The "financial-fallen-angels-VBA" repository contains VBA code that runs a web query on specific stock tickers, retrieves PE ratio data, and determines if a stock is a "Fallen Angel" based on its historical PE ratios. This project is inspired by Ben Graham's criteria for identifying "fallen angels" in financial analysis. To identify fallen angels, he looked for companies whose current P/E ratio was less than half of their highest P/E ratio in the last five years. The idea is that a fallen angel might be a company whose stock is currently undervalued compared to its historical performance. By comparing the current P/E ratio to the five-year high, investors could potentially identify stocks that have fallen out of favor but might have the potential for recovery.

## Table of Contents

- [Installation](#installation)
- [Usage](#usage)
- [Contributing](CONTRIBUTING.md)
- [Code of Conduct](CODE_OF_CONDUCT.md)
- [License](#license)

## Installation

1. Download the repository.
2. Open the `FallenAngel-ExcelFile.xlsm` Excel file, which contains the VBA code.
3. Enable macros in Excel if prompted.
4. Run the VBA code as described in the usage section.

## Usage

1. Open the Excel file.
2. Navigate to the "Fallen Angel" sheet.
3. Run the macro by clicking the "Process Tickers" button on the "Homework" tab of the ribbon.
4. The information for all stock ticker symbols located contiguously below cell B3 on the "Fallen Angel" sheet will be retrieved, and each will be marked with a "Yes" or "No" in the "Fallen Angel" column.

## Contributing

If you would like to contribute to this project, please follow the guidelines outlined in [CONTRIBUTING.md](CONTRIBUTING.md).

## Code of Conduct

This project adheres to a [Code of Conduct](CODE_OF_CONDUCT.md). By participating, you are expected to uphold this code.

## License

This project is licensed under the [Apache License 2.0](LICENSE).
