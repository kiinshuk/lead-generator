# Lead Generator

A simple web app that generates leads for import/export businesses by searching for companies and extracting their contact information into an Excel file.

## Features

- Search for companies by product and country
- Multiple search queries (importer, distributor, wholesale, trader, supplier)
- Extract company name, email, phone from websites
- Auto-download Excel file with leads

## Installation

```bash
npm install
```

## Usage

```bash
node server.js
```

Then open http://localhost:3000

## Tech Stack

- Node.js
- Express
- Axios + Cheerio (web scraping)
- SheetJS (Excel generation)
