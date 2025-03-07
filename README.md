# SheetAI

## Overview

**SheetAI** applies AI prompts to data in Google Sheets, generating results as additional columns for streamlined classification, enrichment, and analysis. It simplifies tasks such as data classification, enrichment, and processing with minimal manual effort.

![Template Screenshot](https://raw.githubusercontent.com/zyxware/SheetAI/refs/heads/main/doc-assets/template-image.png)

## Features

- **AI-Powered Data Processing** – Uses OpenAI to classify and enrich spreadsheet data.
- **Batch Processing** - Supports batch processing of up to 50,000 requests per batch for efficient handling of large datasets.
- **Automated Column Management** – Dynamically creates columns for extracted data.
- **Cost Tracking** – Logs token usage (including cached tokens) and calculates OpenAI API costs accurately.
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
   - **A3**: `DEBUG` → **B3**: `0` *(Set to **`1`** to enable logging, **`0`** to disable)*
   - **A4**: `BATCH_SIZE` → **B4**: `2000` *(Number of rows to process in each batch)*
   - **A5**: `TEMPERATURE` → **B5**: `0` *(Controls randomness: 0 = deterministic, 1 = creative)*
   - **A6**: `MAX_TOKENS` → **B6**: `256` *(Maximum tokens in response)*
   - **A7**: `SEED` → **B7**: `101` *(Seed for reproducible results)*

Checkout OpenAI documentation for more details on the parameters: https://platform.openai.com/docs/api-reference/completions/create

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

1. Open the **Prompts** sheet.
2. The first row should have these headers:
   - **A1**: `Prompt Name`
   - **B1**: `Prompt Text`
   - **C1**: `Model` *(Optional: Defaults to config value if empty)*
   - **D1**: `Active` *(Set to 1 to enable, 0 to disable)*
   - **E1**: `Temperature` *(Optional: Defaults to config value if empty)*
   - **F1**: `Max Tokens` *(Optional: Defaults to config value if empty)*

3. Enter classification or processing prompts in the rows below.

It is posssible to use multiple prompts in the same sheet, just make sure the column names match exactly as they appear in the **Data** sheet. The system will run all the active prompts and add all the results in the new columns.

You should mention the keys in the response that you want to extract in the prompt, the system will add the results in the new columns.

Example:

```
Classify the company based on Description: {{Company Description}} and Industry: {{Industry}}

Return only valid JSON with the following keys:

 Type: B2B or B2C
 AI Services: Yes or No based on whether the company has AI-related services.
```
In this case, the openai will return a json with the keys `Type` and `AI Services`. and the system will add the results in the new columns.

### **Step 4: Prepare Data Sheet**

1. The **Data** sheet should have a header row (Row 1).
2. Add relevant data columns that AI will process.

### **Step 5: Run SheetAI**

All you need to do is create prompts, use the columns you want to include in the prompt as tokens, and click **OpenAI Tools** -> **Run for All Rows**. Wait for the processing to complete and view the results in newly created columns.

You should always run `Run for First 10 Rows` first to test and optimize your prompt.

For large datasets, you can use the batch processing feature:
1. Click **OpenAI Tools** -> **Create Batch**
2. Once the batch is created, click **OpenAI Tools** -> **Check Batch Status** to monitor progress
3. When the batch is complete, click **OpenAI Tools** -> **Check and Process Batch** to process the results

As this script is not yet verified by Google, you should be asked to authorize the script. You should follow the steps:
![Permission to access the AppScript](https://raw.githubusercontent.com/zyxware/SheetAI/refs/heads/main/doc-assets/Permission%20to%20the%20AppScript.png)

We are not capturing any user data or sending data to external systems, the script will run completly on your sheet and only communicate to OpenAI apis. You can get full source code by clicking on Extensions -> AppScript Menu in Google Sheets.

#### **Recommendations:**

- If you have a large dataset, try running `Run for First 10 Rows` first to test and optimize your prompt.
- You can remove all auto-generated columns to rerun on the same data. Auto-generated columns include those created dynamically based on the prompts, such as `Prompt Name - Key` columns, where AI-generated results are stored.
- Once processed, you can copy/export the `Data` sheet for further use or clear the sheet to start fresh with a new dataset.

## Execution Logs & Debugging

- **Execution Log**: Logs prompts sent and responses received.
- **Error Log**: Captures any issues encountered.
- **Cost Summary**: Tracks token usage (including cached tokens) and costs incurred in execution.
- **Batch Status**: Monitors the status of batch processing jobs.

## Troubleshooting

### **Common Issues & Fixes**

| Issue                 | Cause                                 | Fix                                                |
| --------------------- | ------------------------------------- | -------------------------------------------------- |
| Data not writing back | Columns missing or invalid JSON       | Check prompts, Ensure correct column names & valid JSON responses |
| OpenAI API Error      | Invalid API key or quota exceeded     | Verify API key & OpenAI account limits             |
| Batch processing fails| File size too large or timeout        | Reduce batch size in Config sheet                  |

## Support

This application was developed as an internal tool and we would continue to improve and optimize this for as long as we use it. If however you would like us to customize this or build a similar or related system to automate your tasks with AI, we would be available for commercial support.

## About Us

Zyxware Technologies enables brands to define and execute the next steps in their digital transformation journey; a journey towards rich, personalised experiences for their stakeholders. Zyxware assures sustainable results for businesses on the twin engines of privacy centered data strategy and digital services focused on scalability and adaptiveness.

Headquartered in India, with offices in the USA & Australia - Zyxware has a team with competencies in Business, Engineering, and Experience, enabling brands to achieve digital agility and leadership in their categories since 2006.

We specialize in transforming digital operations with AI and Low-Code/No-Code automation. As advocates for Free Software, we contribute through code and community initiatives. Learn more at https://www.zyxware.com 

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


