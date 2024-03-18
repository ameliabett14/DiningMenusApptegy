function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var csvMenuEntries = [{name: "Pick Your School", functionName: "writeCodes"},{name: "Apply Formatting", functionName: "format"}];
  ss.addMenu("Dining Menu Options", csvMenuEntries);
};

function format(){

  applyFormattingWeekOne();
  applyFormattingWeekTwo();

}


function applyFormattingWeekOne() {
  var stylesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var primarycolor = stylesheet.getRange("B26").getValue();
  var secondarycolor = stylesheet.getRange("B27").getValue();

  var targetsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week One');
  var header1 = targetsheet.getRange("B13");
  var header2 = targetsheet.getRange("B24");
  var breakfastdays = targetsheet.getRange("B14:R14");
  var lunchdays = targetsheet.getRange("B26:R26");


  var color1 = "#000000";  // Black
  var color2 = "#FFFFFF";  // White

  function getContrastRatio(color1, color2) {
  // Convert color values to RGB
  var rgb1 = hexToRgb(color1);
  var rgb2 = hexToRgb(color2);

  // Calculate luminance for each color
  var luminance1 = getLuminance(rgb1);
  var luminance2 = getLuminance(rgb2);

  // Calculate contrast ratio
  var contrastRatio = (Math.max(luminance1, luminance2) + 0.05) / (Math.min(luminance1, luminance2) + 0.05);
  return contrastRatio.toFixed(2);
}



function hexToRgb(hex) {
  var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
  hex = hex.replace(shorthandRegex, function (m, r, g, b) {
    return r + r + g + g + b + b;
  });

  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      }
    : null;
}

function getLuminance(rgb) {
  var sRGB = [rgb.r / 255, rgb.g / 255, rgb.b / 255];
  var sRGBAdjusted = sRGB.map(function (channel) {
    if (channel <= 0.03928) {
      return channel / 12.92;
    } else {
      return Math.pow((channel + 0.055) / 1.055, 2.4);
    }
  });

  return 0.2126 * sRGBAdjusted[0] + 0.7152 * sRGBAdjusted[1] + 0.0722 * sRGBAdjusted[2];
}

    ///var color1 = "#000000";  // Black
    ///var color2 = "#FFFFFF";  // White

  var ratiosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  
  var ratio = getContrastRatio(primarycolor, color1);
  ratiosheet.getRange("B30").setValue(ratio);
  var ratio2 = getContrastRatio(primarycolor, color2);
  ratiosheet.getRange("B31").setValue(ratio2);

  var forcedRatio = ratiosheet.getRange("B30").getValue();
  var forcedRatio2 = ratiosheet.getRange("B31").getValue();

  console.log("With Black: ", forcedRatio);

  console.log("With White: ", forcedRatio2);

  if (forcedRatio2 > forcedRatio)
  {
    var primarytext = color2;
  } else {var primarytext = color1;}

  var ratiob = getContrastRatio(secondarycolor, color1);
  ratiosheet.getRange("B34").setValue(ratiob);
  var ratiob2 = getContrastRatio(secondarycolor, color2);
  ratiosheet.getRange("B35").setValue(ratiob2);

  var forcedRatioB = ratiosheet.getRange("B34").getValue();
  var forcedRatioB2 = ratiosheet.getRange("B35").getValue();

  console.log("With Black: ", forcedRatioB);
  console.log("With White: ", forcedRatioB2);

  if (forcedRatioB2 > forcedRatioB)
  {
    var secondarytext = color2;
  } else {var secondarytext = color1;}



  
  // Apply formatting to the cell
  header1.setBackground(primarycolor);
  header1.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  header1.setFontColor(primarytext);
  breakfastdays.setBackground(secondarycolor);
  breakfastdays.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  breakfastdays.setFontColor(secondarytext);

  header2.setBackground(primarycolor);
  header2.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  header2.setFontColor(primarytext);
  lunchdays.setBackground(secondarycolor);
  lunchdays.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  lunchdays.setFontColor(secondarytext);
  

}

function applyFormattingWeekTwo() {
  var stylesheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data');
  var primarycolor = stylesheet.getRange("B26").getValue();
  var secondarycolor = stylesheet.getRange("B27").getValue();

  var targetsheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Two');
  var header1 = targetsheet.getRange("B13");
  var header2 = targetsheet.getRange("B24");
  var breakfastdays = targetsheet.getRange("B14:R14");
  var lunchdays = targetsheet.getRange("B26:R26");


  var color1 = "#000000";  // Black
  var color2 = "#FFFFFF";  // White

  function getContrastRatio(color1, color2) {
  // Convert color values to RGB
  var rgb1 = hexToRgb(color1);
  var rgb2 = hexToRgb(color2);

  // Calculate luminance for each color
  var luminance1 = getLuminance(rgb1);
  var luminance2 = getLuminance(rgb2);

  // Calculate contrast ratio
  var contrastRatio = (Math.max(luminance1, luminance2) + 0.05) / (Math.min(luminance1, luminance2) + 0.05);
  return contrastRatio.toFixed(2);
}



function hexToRgb(hex) {
  var shorthandRegex = /^#?([a-f\d])([a-f\d])([a-f\d])$/i;
  hex = hex.replace(shorthandRegex, function (m, r, g, b) {
    return r + r + g + g + b + b;
  });

  var result = /^#?([a-f\d]{2})([a-f\d]{2})([a-f\d]{2})$/i.exec(hex);
  return result
    ? {
        r: parseInt(result[1], 16),
        g: parseInt(result[2], 16),
        b: parseInt(result[3], 16),
      }
    : null;
}

function getLuminance(rgb) {
  var sRGB = [rgb.r / 255, rgb.g / 255, rgb.b / 255];
  var sRGBAdjusted = sRGB.map(function (channel) {
    if (channel <= 0.03928) {
      return channel / 12.92;
    } else {
      return Math.pow((channel + 0.055) / 1.055, 2.4);
    }
  });

  return 0.2126 * sRGBAdjusted[0] + 0.7152 * sRGBAdjusted[1] + 0.0722 * sRGBAdjusted[2];
}

    ///var color1 = "#000000";  // Black
    ///var color2 = "#FFFFFF";  // White

var ratiosheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Data");
  
  var ratio = getContrastRatio(primarycolor, color1);
  ratiosheet.getRange("B38").setValue(ratio);
  var ratio2 = getContrastRatio(primarycolor, color2);
  ratiosheet.getRange("B39").setValue(ratio2);

  var forcedRatio = ratiosheet.getRange("B38").getValue();
  var forcedRatio2 = ratiosheet.getRange("B39").getValue();

  console.log("With Black: ", forcedRatio);

  console.log("With White: ", forcedRatio2);

  if (forcedRatio2 > forcedRatio)
  {
    var primarytext = color2;
  } else {var primarytext = color1;}

  var ratiob = getContrastRatio(secondarycolor, color1);
  ratiosheet.getRange("B42").setValue(ratiob);
  var ratiob2 = getContrastRatio(secondarycolor, color2);
  ratiosheet.getRange("B43").setValue(ratiob2);

  var forcedRatioB = ratiosheet.getRange("B42").getValue();
  var forcedRatioB2 = ratiosheet.getRange("B43").getValue();

  console.log("With Black: ", forcedRatioB);
  console.log("With White: ", forcedRatioB2);

  if (forcedRatioB2 > forcedRatioB)
  {
    var secondarytext = color2;
  } else {var secondarytext = color1;}



  
  // Apply formatting to the cell
  header1.setBackground(primarycolor);
  header1.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  header1.setFontColor(primarytext);
  breakfastdays.setBackground(secondarycolor);
  breakfastdays.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  breakfastdays.setFontColor(secondarytext);

  header2.setBackground(primarycolor);
  header2.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
  header2.setFontColor(primarytext);
  lunchdays.setBackground(secondarycolor);
  lunchdays.setFontColor(secondarytext);
  lunchdays.setBorder(true, true, true, true, true, true, "#000000", SpreadsheetApp.BorderStyle.SOLID);
}

function writeCodes(){
  var ui = SpreadsheetApp.getUi();
  var thisweek = ui.prompt("What is the URL for your Dining Page?").getResponseText();
  var nextweek = ui.prompt("What is the URL for your next week Dining Page?").getResponseText();
  

  var formulasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('data');

  var breakfast1 = formulasheet.getRange("A1");
  var breakfast2 = formulasheet.getRange("A6");
  var breakfastdate1 = formulasheet.getRange("C1");
  var breakfastdate2 = formulasheet.getRange("C6");
  var lunch1 = formulasheet.getRange("A12");
  var lunch2 = formulasheet.getRange("A17");
  var lunchdate1 = formulasheet.getRange("C12");
  var lunchdate2 = formulasheet.getRange("C17");
  var logo = formulasheet.getRange("B23");
  var style = formulasheet.getRange("B25");

  ///=IMPORTXML("02WEB", "//div[@class='breakfast']")
  ///=INDEX(IMPORTXML("https://bellaar.sites.thrillshare.com/o/amelia/dining", "//div[@class='logo']//img/@src"), 1)

  breakfast1.setValue('=IMPORTXML("' + thisweek + '","//div[@class=' + "'breakfast'" + ']")');
  lunch1.setValue('=IMPORTXML("' + thisweek + '","//div[@class=' + "'lunch'" + ']")');
  breakfast2.setValue('=IMPORTXML("' + nextweek + '","//div[@class=' + "'breakfast'" + ']")');
  lunch2.setValue('=IMPORTXML("' + nextweek + '","//div[@class=' + "'lunch'" + ']")');
  breakfastdate1.setValue('=IMPORTXML("' + thisweek + '","//div[@class=' + "'date'" + ']")');
  breakfastdate2.setValue('=IMPORTXML("' + nextweek + '","//div[@class=' + "'date'" + ']")');
  lunchdate1.setValue('=IMPORTXML("' + thisweek + '","//div[@class=' + "'date'" + ']")');
  lunchdate2.setValue('=IMPORTXML("' + nextweek + '","//div[@class=' + "'date'" + ']")');
  style.setValue('=INDEX(IMPORTXML("' + thisweek + '", "//@style"), 1)');
  logo.setValue('=INDEX(IMPORTXML("' + thisweek + '", "//div[@class='+"'logo'"+']//img/@src"), 1)');

  applyFormattingWeekOne();
  applyFormattingWeekTwo();
}
