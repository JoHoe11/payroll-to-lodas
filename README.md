# Payroll-to-Lodas
[![npm version](https://img.shields.io/npm/v/papaya)](https://www.npmjs.com/package/papaya)

Takes payroll-related files (Lohnjournal, Personalkosten, payslips) and converts them into a LODAS-compatible input file. This is a web app so it's easy to setup. It saves alot of manual labor. The Inputs are the current month so the function knows which Code is NBA or not. The special codes are not requierd but can be used to exclude specified codes.


## Features

- Converts payroll-related files (Lohnjournal, Personalkosten, payslips) to LODAS-compatible input.
- Supports processing of Lohnjournal files.
- Supports processing of Personalkosten files.
- Supports processing of payslips files.

## Tech Stack / Key Dependencies

- JavaScript
- CSS
- HTML
- Vite
- Sass
- date-fns

## File Structure Overview

```text
.
├── README.md
├── dornbach.css
├── help.js
├── index.html
├── lib
├── package.json
├── script.js
├── styles.css
├── ui.js
├── vite.config.js
```

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/JoHoe11/payroll-to-lodas.git
   cd payroll-to-lodas
   ```
2. Install the dependencies:
   ```bash
   npm install
   ```

## Usage / Getting Started

1.  To start the development server:
    ```bash
    npm run dev
    ```
2. To build the project for production:
   ```bash
   npm run build
   ```
3. To preview the built project:
    ```bash
    npm run preview
    ```



## Contributing

Pull requests are welcome. 


## License

Distributed under the MIT License. See `LICENSE` file for more information.

## Author/Acknowledgements

Author: JoHoe11
