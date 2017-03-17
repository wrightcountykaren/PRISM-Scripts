// TODO: this isn't working in IE??
// TODO: add "about scripts" page
// TODO: add footer to all docs
// TODO: pretty up the html to view script title in a single header (this should borrow from the alpha split in FuncLib)
// TODO: set to evaluate master and not release
// TODO: incorporate date.js and determine dates in a cleaner way, with filtering and formatting custom to the user (last 3 months, etc)
// TODO: add a feature to switch branches (master or insert your own)
// TODO: replace warning text with a spinner or something so as not to alarm folks with slow connections, move warning text to something that happens if connection not made
// TODO: create expand all functionality

function displayChangelogInfo() {
    
    var listOfScriptsHTML = document.getElementById("changelogContents")
    
    var functionToCheckFor = "changelog_update"
    
    // read text from URL location to get the list of scripts
    var request = new XMLHttpRequest();
    request.open('GET', 'https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/~complete-list-of-scripts.vbs', false);
    
    // This sends the request for info and does all of the hard work
    request.onreadystatechange = function () {
        // If the data is there, then...
        if (request.readyState === 4 && request.status === 200) {
            // create a new variable called "type" which handles the response header, or "type of content we're dealing with"
            var type = request.getResponseHeader('Content-Type');
            
            // If it's text, that means it's probably working and we can proceed!
            if (type.indexOf("text") !== 1) {
                
                // Create a variable filled with the contents of the FuncLib file
                var listOfScriptsArray = request.responseText.split("\n");
                
                var listOfScripts = "";
                
                
                for (var i = 0; i < listOfScriptsArray.length; i++) {
                    if (listOfScriptsArray[i].startsWith("cs_scripts_array(script_num).script_name")) {
                        
                        // Creating a friendly name for the new script
                        var scriptFriendlyName = listOfScriptsArray[i].replace('cs_scripts_array(script_num).script_name', '').replace(/"/g, '').replace('=', '').trim();
                        
                        // Getting the category, which is always on the next line
                        var scriptCategory = listOfScriptsArray[i + 1].slice((listOfScriptsArray[i + 1].length - listOfScriptsArray[i + 1].lastIndexOf("=")) * -1).replace(/"/g, '').replace('=', '').trim();
                        
                        // Getting the URL for the script file
                        var scriptURL = 'https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/' + scriptCategory + '\\' + scriptFriendlyName.toLowerCase().replace(/ /g, '-') + '.vbs';
                                                
                        // read text from URL location to get the list of scripts
                        var scriptCheck = new XMLHttpRequest();
                            
                        // This sends the request for info and does all of the hard work
                        scriptCheck.onreadystatechange = function () {
                            // If the data is there, then...
                            if (scriptCheck.readyState === 4 && scriptCheck.status === 200) {
                                
                                // Gets the whole file
                                var data = scriptCheck.responseText;
                                
                                // Splits into array
                                var fileArray = data.split("\n");
                                
                                // This is a regex that checks for changelog_update
                                var re = new RegExp(functionToCheckFor, "i");

                                // Goes through the array created above, and checks for the changelog_update elements.
                                for (var j = 0; j < fileArray.length; j++) {
                                    if (fileArray[j].search(re) != -1) {
                                        
                                        // Regex for the changelog display, which is case insensitive
                                        var regexForChangelogDisplay = new RegExp("changelog_display", "i");
                                        
                                        // Escapes if the line meets certain criteria: 
                                        //      - it's an example
                                        //      - the string "changelog_display" is found (which means the end of the changelog block is here)
                                        if (fileArray[j] == "\'Example: call changelog_update(\"01/01/2000\", \"The script has been updated to fix a typo on the initial dialog.\", \"Jane Public, Oak County\")") {
                                            continue;
                                        } else if (fileArray[j].search(regexForChangelogDisplay) != -1) {
                                            break;
                                        }
                                        
                                        // Determines date, text, and scriptwriter
                                        var changelogEntryArray = fileArray[j].split('\"');         // splits into an array
                                        var changelogEntryDate = new Date(changelogEntryArray[1]);  // item [1] is the entry date
                                        var changelogEntryText = changelogEntryArray[3];            // item [3] is the entry text
                                        var changelogEntryScriptwriter = changelogEntryArray[5];    // item [5] is the entry scriptwriter
                                        
                                        var today = new Date();                            
                                        var changelogDateDiff = parseInt((today - changelogEntryDate)/(1000*60*60*24));
                                        
                                        if (changelogDateDiff <= 30) {
                                            // This is the part that writes to the HTML doc
                                            listOfScriptsHTML.insertAdjacentHTML('beforeend', 
                                            
                                            "<h4><a href=\'" + scriptURL + "\' target=\'_blank\'>" + scriptCategory.toUpperCase() + " - " + scriptFriendlyName + "</a></h4> \n" + 
                                            "<h5>" + changelogEntryDate.toDateString() + "</h5> \n" + 
                                            "<p>" + changelogEntryText + "</p> \n" + 
                                            "<p><strong> Completed by " + changelogEntryScriptwriter + ". </strong></p>"
                                            );
                                        }
                                    };
                                }                                
                            } 
                        }
                        scriptCheck.open('GET', scriptURL, false);                            
                        scriptCheck.send();
                    }
                }
            }
        }
    }
    request.send(null);
}
