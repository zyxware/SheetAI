# SheetAI

## Overview

**SheetAI** applies AI prompts to data in Google Sheets, generating results as additional columns for streamlined classification, enrichment, and analysis. It simplifies tasks such as data classification, enrichment, and processing with minimal manual effort.

![Template Screenshot](https://raw.githubusercontent.com/zyxware/SheetAI/refs/heads/main/doc-assets/template-image.png)

## Features

- **AI-Powered Data Processing** – Uses OpenAI to classify and enrich spreadsheet data.
- **Automated Column Management** – Dynamically creates columns for extracted data.
- **Cost Tracking** – Logs token usage and calculates OpenAI API costs.
- **Debugging & Logging** – Execution logs for tracking inputs, outputs, and errors.
- **Customizable Prompts** – Define prompts and AI models per task.

## Use Cases

As this is a generic framework to apply prompts to data in a Google Sheet, the use cases are unlimited. Some examples include:

- **Extracting Information from Unstructured Text** – Extract key services from company descriptions.
- **Job Title Classification** – Categorize people based on their designations.
- **Address Parsing** – Extract city, state, and country from a single text field.
- **Lead Prioritization from LinkedIn Messages** – Identify potential contacts based on past engagement.
- **Sentiment Analysis** – Analyze customer feedback and classify sentiment.
- **Entity Extraction** – Extract key details such as names, dates, and locations.
- **Lead Qualification** – Process sales lead data to determine qualification status.
- **Content Summarization** – Generate short summaries from large text inputs.
- **Spam Detection** – Identify spam messages from customer inquiries.

## Setup Instructions

### **System Requirements**

- A Google Account
- Use Template to [make a copy of the sheet for your use](https://docs.google.com/spreadsheets/d/1FF_uPaxJe3_8MCA_UdPq9xn64yOyHRfyopKeyJ6uWUA/template/preview). 
- An OpenAI API Key - You can create one [here](https://platform.openai.com/api-keys)


### **Step 1: Setup the Google Sheet**

The template has all the structure ready, the data sheet is where you put your data.

### **Step 2: Configure API Key & Settings**

1. Open your Google Sheet.
2. Locate the **Config** sheet (already included in the setup).
3. Configure the following values:
   - **A1**: `API_KEY`  → **B1**: *(Your OpenAI API Key)*
   - **A2**: `DEFAULT_MODEL` → **B2**: `gpt-4o-mini` *(or another model)*
   - **A3**: `DEBUG_MODE` → **B3**: `on` *(Set to **``** to enable logging, **`off`** to disable)*

### **Step 3: Define Prompts**

#### **Using Tokens in Prompts**

Tokens are placeholders that represent column values in your Google Sheet. You can include tokens in your prompt using double curly braces `{{ }}` around the column name.

For example, if your data sheet has columns **Company Description** and **Industry**, you can create a prompt like this:

```
Classify the company based on Description: {{Company Description}} and Industry: {{Industry}}

Return only valid JSON with the following keys:

 Type: B2B or B2C
 AI Services: Yes or No based on whether the company has AI-related services.
```

SheetAI will replace `{{Company Description}}` and `{{Industry}}` with the actual values from the corresponding columns in each row when sending the prompt to OpenAI. Ensure that the column names match exactly as they appear in the **Data** sheet. For example:

- ✅ Correct: `{{Company Description}}`
- ❌ Incorrect: `{{company_description}}` (case-sensitive) or `{{CompanyDescription}}` (missing space)

If column names contain spaces, they must be written exactly as they appear in the headers of the **Data** sheet.

When writing prompts, make sure the column names match exactly as they appear in the **Data** sheet. If the column name contains spaces, ensure they are written correctly within `{{ }}` in the prompt.

1. Open the **Prompts** sheet.
2. The first row should have these headers:
   - **A1**: `Prompt Name`
   - **B1**: `Prompt Text`
   - **C1**: `Model` *(Optional: Defaults to **``** if empty)*
3. Enter classification or processing prompts in the rows below.

Example:

```
Classify the company based on Description: {{Company Description}} and Industry: {{Industry}}

Return only valid JSON with the following keys:

 Type: B2B or B2C
 AI Services: Yes or No based on whether the company has AI-related services.
```

### **Step 4: Prepare Data Sheet**

1. The **Data** sheet should have a header row (Row 1).
2. Add relevant data columns that AI will process.

### **Step 5: Run SheetAI**

All you need to do is create prompts, use the columns you want to include in the prompt as tokens, and click **OpenAI Tools** -> **Run for All Rows**. Wait for the processing to complete and view the results in newly created columns.

As this script is not yet verified by Google, you should be asked for permission. You should follow the steps:
![Permission to access the AppScript](https://raw.githubusercontent.com/zyxware/SheetAI/refs/heads/main/doc-assets/Permission%20to%20the%20AppScript.png)

We are not capturing any user data or sending data to external systems, the script will run completly on your sheet and only communicate to OpenAI apis. You can get full source code by clicking on Extensions -> AppScript Menu in Google Sheets.

#### **Recommendations:**

- If you have a large dataset, try running `Run for First 10 Rows` first to test and optimize your prompt.
- You can remove all auto-generated columns to rerun on the same data. Auto-generated columns include those created dynamically based on the prompts, such as `Prompt Name - Key` columns, where AI-generated results are stored.
- Once processed, you can copy/export the `Data` sheet for further use or clear the sheet to start fresh with a new dataset.

## Execution Logs & Debugging

- **Execution Log**: Logs prompts sent and responses received.
- **Error Log**: Captures any issues encountered.
- **Cost Summary**: Tracks the cost incured in execution.

## Troubleshooting

### **Common Issues & Fixes**

| Issue                 | Cause                                 | Fix                                                |
| --------------------- | ------------------------------------- | -------------------------------------------------- |
| Data not writing back | Columns missing or invalid JSON       | Check prompts, Ensure correct column names & valid JSON responses |
| OpenAI API Error      | Invalid API key or quota exceeded     | Verify API key & OpenAI account limits             |

## Future Improvements

- Support for **batch API processing**, which would allow processing multiple rows in parallel, reducing execution time and improving efficiency for large datasets..

## Support

This application was developed as an internal tool and we would continue to improve and optimize this for as long as we use it. If however you would like us to customize this orbuild a similar or related system to automate your tasks with AI, we would be available for commercial support.

## About Us

We are a software development company. We help organizations transforms their digital operations. We love Free Software and try to contribute to it both through direct code contributions as well as through community level promotional activities. You can see more about us at https://www.zyxware.com

## Contact
https://www.zyxware.com/contact-us

## Source Repository

https://github.com/zyxware/SheetAI

## Reporting Issues

https://github.com/zyxware/SheetAI/issues

## License

GPL v2 – Free to use & modify.

## Need Help or Commercial Support?

If you have any questions, feel free to [contact us](https://www.zyxware.com/contact-us)


