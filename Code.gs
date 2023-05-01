function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('AI Functions Sidebar')
      .addItem('Show Sidebar', 'showForm')
      .addToUi();
}

function showForm() {
  var htmlOutput = HtmlService.createHtmlOutputFromFile('form.html')
      .setTitle('AI Functions');
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

currentCell = SpreadsheetApp.getActiveSheet().getActiveCell().getValue();
keyCell = SpreadsheetApp.getActive().getSheetByName('KeySheet').getRange(2,2).getValue();
console.log(keyCell)
const SECRET_KEY = keyCell;
const MAX_DAVINCI_TOKENS = 2000;

/**
 * Usage =Bulletize(text,precontext,postcontext)
 * If no precontext or postcontext is added, then the default
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Bulletize(model,text, precontext, postcontext){
  var defaultcontext = "Create bullets out of this text. Here is the text: " 
  var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " return only the bulleted items and no other text";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =Summarize(text,precontext,postcontext)
 * If no precontext or postcontext is added, then the default
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Summarize(model, text, precontext, postcontext){
  var defaultcontext = "Summarize the following text concisely. Here is the text: " 
  var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " ";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =Detail(text,precontext,postcontext)
 * If no precontext or postcontext is added, then the default
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Detail(model, text, precontext, postcontext){
  var defaultcontext = "Add more detail to the following text. Here is the text: "
  var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " ";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =Sentiment(model,text,precontext,postcontext)
 * Determines sentiment
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Sentiment(model,text, precontext, postcontext){
  var defaultcontext = "Determine the sentiment of this text: " 
  var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " return only the sentiment as a single word response and no other text";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =Categorize(model,text,precontext,postcontext)
 * Determines categories from a provided list of categories as a comma seperated list
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Categorize(model,text, categories){
  var cats = getTableData(SpreadsheetApp.getActiveSheet(),categories)
  var defaultcontext = "Determine which category this text best belongs to from this list: " +  cats
  var text = text

    precontext = defaultcontext;
    postcontext = " return only the category as a single word response and no other text";

    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =FormulaHelper(model,text,precontext,postcontext, explain)
 * Builds formulas for google sheets, set explain to true to get an explanation
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function FormulaHelper(model,text, explain=false){
  var defaultcontext = "Create a formula for google sheets that : " 
  var text = text
  var explain = explain
  var postcontext = " return only the suggested formula and no other text either before or after the formula";

  if(explain == true){
     var postcontext = "";
  }

    var prompt = defaultcontext + text + postcontext
    
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}
/**
 * Usage =Expand(model,text,precontext,postcontext)
 * If no precontext or postcontext is added, then the default
 * @param {text,"This is prepended","This is appended at the end"}
 * @returns  Bullets to the cell you call it from
 * @customfunction
 */
function Expand(model, text, precontext, postcontext){
  var defaultcontext = "Expand on this text. Here is the text: " 
  var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " ";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}

/**
 * Usage =Analyze(text,precontext,postcontext)
 * Can draw conclusions from structured data
 * @param {range for the structured data}
 * @returns  Insights
 * @customfunction
 */
function Analyze(model, text, precontext, postcontext){
  var defaultcontext = "Analyze the data in the table:  " 
  var text = getTableData(SpreadsheetApp.getActiveSheet(),text);

    if((precontext == "" | precontext == null | precontext == "undefined" | precontext == "None")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " return only the results of the analysis and no other text";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}

/**
 * Usage =Direct(model,text,precontext,postcontext)
 * Call models directly, use CONCATENATE to chain together inputs
 * @param {model, text}
 * @returns  Insights
 * @customfunction
 */
function Direct(model, text, precontext, postcontext){
    var defaultcontext = "  " 
    var text = text
    if((precontext == "" | precontext == null | precontext == "undefined")&&( postcontext == "" | postcontext == null |postcontext == "undefined")){
    precontext = defaultcontext;
    postcontext = " ";
  } else if (precontext == "" | precontext == null | precontext == "undefined") {
    precontext = defaultcontext;
  } else if ( postcontext == "" | postcontext == null |postcontext == "undefined"){
    postcontext = " ";
  }
    var prompt = precontext + text + postcontext
    if(model == "Davinci3"){
      return Davinci3(prompt);
    } else if (model == "Turbo"){
      return TURBO(prompt);
    } else if (model == "GPT4"){
      return GPT4(prompt);
    } else {
      return "No model selected";
    }
}

function createChart() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange('A8:L18'); // Adjust the range to include your data
  
  // Create the chart
  var chartBuilder = sheet.newChart();
  chartBuilder
    .addRange(dataRange)
    .setChartType(Charts.ChartType.LINE)
    .setOption('title', 'Sales by Month')
    .setPosition(2, 3, 0, 0);
  
  // Add the chart to the sheet
  sheet.insertChart(chartBuilder.build());
}

function GPT4(prompt, model="gpt-4", temperature=0.5) {
     let apiKey = SECRET_KEY;
     let endpoint = "https://api.openai.com/v1/chat/completions";
     let payload = {
       messages:[
        {role: 'system', content: 'This is where you describe what gpt does and how it answers.'},
        {role: 'user', content: 'This is an example question'},
        {role: 'assistant', content: "This is an example answer"},
        {role: 'user', content: `${prompt}`},
            ],
       model: model,
       temperature: temperature,
       max_tokens: MAX_TURBO_TOKENS,
     };
     let options = {
       method: "post",
       contentType: "application/json",
       payload: JSON.stringify(payload),
       headers: {
         "Authorization": "Bearer " + apiKey
       }
     };
     
     let response = UrlFetchApp.fetch(endpoint, options);
     let json = response.getContentText();
     let data = JSON.parse(json);    
     return data.choices[0].message.content.trim();
}

function Davinci3(prompt, model="text-davinci-003", temperature=0.5) {
    let endpoint = "https://api.openai.com/v1/completions";
    let apiKey = SECRET_KEY;
    let payload = {
        model: model,
        prompt: prompt,
        max_tokens: MAX_DAVINCI_TOKENS,
        temperature: temperature
    };
     
    let options = {
       method: "post",
       contentType: "application/json",
       payload: JSON.stringify(payload),
       headers: {"Authorization": "Bearer " + apiKey}
    };
     
    let response = UrlFetchApp.fetch(endpoint, options);
    let json = response.getContentText();
    let data = JSON.parse(json);
     
    return data.choices[0].text.trim();
}

const MAX_TURBO_TOKENS = 3000;
   
function TURBO(prompt, model="gpt-3.5-turbo", temperature=0.5) {
     let apiKey = SECRET_KEY;
     let endpoint = "https://api.openai.com/v1/chat/completions";
     let payload = {
       messages:[
        {role: 'system', content: 'This is where you describe what gpt does and how it answers.'},
        {role: 'user', content: 'This is an example question'},
        {role: 'assistant', content: "This is an example answer"},
        {role: 'user', content: `${prompt}`},
            ],
       model: model,
       temperature: temperature,
       max_tokens: MAX_TURBO_TOKENS,
     };
     let options = {
       method: "post",
       contentType: "application/json",
       payload: JSON.stringify(payload),
       headers: {
         "Authorization": "Bearer " + apiKey
       }
     };
     
     let response = UrlFetchApp.fetch(endpoint, options);
     let json = response.getContentText();
     let data = JSON.parse(json);    
     return data.choices[0].message.content.trim();
}


function getTableData(sheet, range) {
  var datacells = sheet.getRange(range).getValues();
  var table = '';

  datacells.forEach(function (row) {
    table += '| ' + row.join(' | ') + ' |\n';
  });
  return table;
}
function callFunction(data2) {
    var functionName = data2.functionName;
    var textCell = data2.text;
    var precontextCell = data2.precontext;
    var postcontextCell = data2.postcontext;
    var outputCell = data2.cell;
    var model = data2.modelName;
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var text = sheet.getRange(textCell).getValue();
    var precontext = precontextCell;
    var postcontext = postcontextCell;

    var range = sheet.getRange(outputCell);

   if (functionName === 'Bulletize') {
        var formula = '=Bulletize("'+model+'","'+text+'","'+precontext+'","'+postcontext+'")';
    } else if (functionName === 'Summarize') {
        var formula = '=Summarize("'+model+'","'+text+'","'+precontext+'","'+postcontext+'")';
    } else if (functionName === 'Detail') {
        var formula = '=Detail("'+model+'","'+text+'","'+precontext+'","'+postcontext+'")';
    } else if (functionName === 'Expand') {
        var formula = '=Expand("'+model+'","'+text+'","'+precontext+'","'+postcontext+'")';
    } else if (functionName === 'Direct') {
        var formula = '=Direct("'+model+'","'+text+'","'+precontext+'","'+postcontext+'")';
    } else if (functionName === 'Analyze') {
        var formula = '=Analyze("'+model+'","'+textCell+'","'+precontext+'","'+postcontext+'")';
    } else {formula = "'=No formula set"}

    range.setFormula(formula);
  }
