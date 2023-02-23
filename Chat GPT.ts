async function main(workbook: ExcelScript.Workbook) {
    
    // Top level constants & variables.
    
    const apiKey = workbook.getWorksheet("API").getRange("B1").getValue();
    const endPoint: string = "https://api.openai.com/v1/completions";
    const sheet = workbook.getWorksheet("Data");
    const productsCount = 13;
    const keywordsCount = 20;
    const productNamesRange = `B1:B${productsCount}`;
    const keywordsRange = `C1:C${productsCount}`;
    const promptStart = "Use this product name \"";
    const promptEnd = `\" and generate ${keywordsCount} unique SEO Keywords for it.`;

    const productNames = sheet.getRange(productNamesRange).getValues();
    const model: string = "text-davinci-002";

    sheet.getRange(keywordsRange).setValue("");

    for (let i = 0; i < productsCount; i++){
        let productName = productNames[i].toString().replace("|", "");
        
        if (productName === ""){
            continue;
        }
        
        let query = promptStart;
        
        query += productName;
        query += promptEnd;

        const prompt: (string | boolean | number) = query;

        // Set the HTTP Headers

        const headers: Headers = new Headers();
        headers.append("Content-Type", "application/json");
        headers.append("Authorization", `Bearer ${apiKey}`);

        // Set the HTTP Request Body.
        
        const body: (string | boolean | number) = JSON.stringify({
            model: model,
            prompt: prompt,
            max_tokens: 2048,
            n: 1,
            temperature: 1,
        });

        // Get the response from ChatGPT.
        
        const response: Response = await fetch(
            endPoint, {
                method: "POST",
                headers: headers,
                body: body,
            }
        );

        const json: { choices: {text: (string | boolean | number)}[]} = await response.json();
        const text: (string | boolean | number) = json.choices[0].text.toString();

        // Get the cell in which we intend to store our data.
        
        const output = sheet.getRange(`C${i+1}`);
        
        // Clean the data & store it.
        
        output.setValue(cleanData(text));
    }
}

function cleanData(data: string){
    // Get each item of the response by using the new line as a separator.
    const listItems = data.split("\n");
    
    // Remove potentially unwanted keywords & special characters.
    const regex = /\b(keyword(s)?|digital marketing|seo|\\|\/|\||\(|\)|\$|%|\^|&|\*|@|!|"|'|`)\b/gi;
    const legalKeywords = listItems.map((item) =>{
        return item.replace(regex, "");
    });
    
    // Get a unique list of responses.
    const uniqueStringList = Array.from(new Set(legalKeywords));
    
    // Remove leading numbers from each response.
    const listItemsWithoutNumbers = uniqueStringList.map((item) =>{
        return item.replace(/^\d+\.\s+/, "");
    });
    
    // Return a comma-separated keyword string.
    return listItemsWithoutNumbers.join(",");
}
