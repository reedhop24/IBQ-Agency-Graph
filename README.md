# IBQ-Agency-Graph

#### This project was built using Node.Js, Express, Chartist, XLSX, and the Node File System module. 

#### The agency graph is an endpoint I built for IBQ, the need for this came about due to they way we were sending User Statistics to our agents. If an agent wanted access their agencies report (i.e. quotes, bind, and pay done broken down by Agent) we were just emailing them a PDF with all of that information scattered about. Since our UI can run HTML, an endpoint that consumes an Excel spreadsheet and returns HTML seemed like a good choice. 

#### We did not have to alter the code that generates and sends the Excel spreadsheet other than changing the endpoint. To generate the charts that represents the spreadsheet I am using Chartist. Chartist takes data for the X Y axis along with the actual user data and generates HTML which forms the graph. So my code takes in the spreadsheet, converts it to JSON, loops through the JSON and fills in the chart with the data it collected. 

#### One snag I ran into was that the chart relies on CSS that chartist provides, which normally wouldn't be a problem. However, I needed the response from this API to be one file. So I had the CSS sit in an HTML file wrapped in a style tag then I used the File System module to create a variable with all that CSS saved in it. I then just appended the data from the chartist function on to that variable and return the variable as the response.
