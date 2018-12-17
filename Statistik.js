// skal køres i konsollen (F12) på en side med jQuery, fx https://skat.sharepoint.com/search/skmsrc/Sider/teamwork.aspx 

var siteUrl = "https://skat.sharepoint.com";
var sti = "*";
var emneOrd = "Medarbejdere";
var statisticsType = "ViewsLifeTimeUniqueUsers"; // kan være ViewsLifeTime, ViewsLifeTimeUniqueUsers, ViewsRecent, ViewsRecentUniqueUsers, 
// ViewsLast1Days -> ViewsLast7Days, ViewsLast1DaysUnique -> ViewsLast7DaysUnique, ViewsLastMonths1 -> ViewsLastMonths3, ViewLastMonths1Unique -> ViewsLastMonths3Unique kan ikke sorteres!

var fileType = "aspx";

$.ajax({
  url: siteUrl + "/_api/search/query?querytext='path:" + sti + " Emneord=" + emneOrd + "'&rowlimit=500&selectproperties='Title," + statisticsType + ",path'&sortlist='"+ statisticsType + ":descending'&clienttype='ContentSearchRegular'",
  type: "GET",
  headers: { "accept": "application/json;odata=verbose" },
  success: function(data, status){
      
      var queryAmount = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results.length;
      var globalsitearray = [];
      var globalsitearrayScreenPrint = [];
      for (var a = 0; a < queryAmount; a++) {
            //debugger;
            var pageTitle = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results[a].Cells.results[2].Value;
            var pageCount = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results[a].Cells.results[3].Value;
            var pageUrl = data.d.query.PrimaryQueryResult.RelevantResults.Table.Rows.results[a].Cells.results[4].Value;
            globalsitearray.push(pageTitle + ": "+ pageCount + " (" + pageUrl +")"); // sætter indholdet ind i et array
            var singleStringLine = String(pageTitle + ","+ pageCount + "," + pageUrl +"<br>"); // bygger hvert resultat som en enkelt String
            globalsitearrayScreenPrint.push(singleStringLine); // sætter hver linje ind i et array
        }
    
    console.log("Du har søgt på " + statisticsType + " på: " + sti + ". Der er fundet: " + queryAmount + " resultater.");
       
    console.log(globalsitearray); // skriver resultaterne til konsollen
    //$(".cmsSPO2013").html(globalsitearrayScreenPrint); //skriver listen ud på skærmen, så den kan lægges ind i Excel; tekst til kolonner, seperator er komma
 },
  error: function(err){
    console.log(err);
  }
});

