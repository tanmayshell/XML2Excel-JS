XML to Excel Using JS
Proof Of Concept

PROCEDURE-
• XML to JSON
• JSON to Excel

VERSIONS-
• NodeJS – 18.14.0 (LTS)
• NPM – 9.3.1

MODULES/LIBRARIES USED-
• xml2js – To parse the XML.
• xlsx – To create the Excel file.
• npm install xml2js xlsx

ES6 (ECMAScript 2015)-
• We use ‘import’ statements to import modules instead of using ‘require’ as Modules are new ES6 features and are more compatible whereas ‘require’ is outdated.

• Using require –
const fs = require(‘fs’);
const xml2js = require(‘xml2js’ );
const XLSX = require(‘xlsx’);

• Using modules –
import fs from ‘fs/promises;
import xml2js from ‘xml2js’;
import XLSX from ‘xlsx’;

• We use the ‘async/await’ syntax for asynchronous operations which can be used for-

1. Reading the XML file.
2. Writing the Excel file.

• We use const and let keywords and can use arrow functions in ES6.

NPM
• Node Package Manager is used for managing packages, dependencies, scripts and versions.
• package.json - This file contains the metadata about the project.
• npm init - This command initializes a new NodeJS project and creates a ‘package.json’ file.
• YARN is an alternative for NPM. It is also a package manager and is faster than NPM as it performs more effective concurrent installations.

OTHER LIBRARIES
• Why xml2js?

1. It has a large user base and active community and is also stable and well maintained.
2. This library also supports asynchronous operations.

• Alternatives to ‘xml2js’ library-

1. xml-js (Performance is a concern while dealing with large XML documents)
2. xml2json (Doesn’t provide built-in support for XML schema validation)
3. fast-xml-parser (Performance is a concern while dealing with large XML documents)

• Why xlsx?

1. It has an active community and is simple to use.
2. It offers both client-side and server-side compatibility and thus can be used on the browser(client side) and NodeJS (server-side).

• Alternatives to ‘xlsx’ library-

1. Exceljs (Smaller community and user base)
2. SheetJS/js-xlsx (Integration challenges)

CODING BEST PRACTICES
• Use camelCase for variable and function names.
• Use proper indentation (Prettier extension can be used for the same).
• Add comments for references.
• Modularize code for reusability and maintainability.
• Implement error handling using try and catch blocks to handle errors and exceptions.
• Optimize loops and avoid excessive and unwanted nesting.

COHESION and COUPLING
• Used to assess and improve the modularity and maintainability of the code.

• COHESION (Desired - HIGH)

1. How well a module’s internal elements are related to each other.
2. internal elements may include functions, classes etc.
3. High cohesion makes code easier to understand, test and maintain.

• COUPLING (Desired – LOW)

1. Measures the degree of interdependence between modules or components.
2. Low coupling means that modules are relatively independent and can be modified without affecting other modules.

XML VALIDATION
• It helps to ensure that the XML data is structured correctly and contains the expected elements and attributes.
• ‘xmldom’ library can be used to parse and validate an XML document against XML Schema Definition.
