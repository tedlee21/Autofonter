/**
 * @OnlyCurrentDoc
 *
 * The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */
var fontFam = 'Arial';
var fontSize = 11;

function onOpen(e) {
  var ui = DocumentApp.getUi()

  var menu = ui.createMenu('Auto-Font')
  menu.addItem('Reformat Document Fonts', 'handleEditEvent')
  var styleMenu = ui.createMenu('Style')
  styleMenu.addItem('Times', 'setTimes')
  menu.addSubMenu(styleMenu)
  menu.addToUi()
}

function onInstall(e) {
  onOpen(e);
}

function handleEditEvent() {
  // Get the active document and document length
  var doc = DocumentApp.getActiveDocument().getBody().editAsText();
  var end = doc.getText().length - 1;

  const style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = fontSize ;
  style[DocumentApp.Attribute.FONT_FAMILY] = fontFam;
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

  // Reformat document font settings with specified values
  // var style = getMostUsedFontSettings()
  doc.setAttributes(0, end, style)
}

function getMostUsedFontSettings() {
  // Get the active document's text
  var doc = DocumentApp.getActiveDocument().getBody();
  var text = doc.getText();

  // Split the text into words
  var words = text.split(/\s+/);

  // Create an object to store the font and size counts
  var fontCounts = {};
  var fontSize = {};

  // Iterate over each word and count the fonts
  words.forEach(function(word) {
    word = word.replace(/[^a-zA-Z0-9]/g, "");
    if (word != "" && doc.findText(word) != null) {
      var font = doc.findText(word).getElement().getFontFamily();
      var size = doc.findText(word).getElement().getFontSize();
      if (fontCounts[font]) {
        fontCounts[font]++;
      } else {
        fontCounts[font] = 1;
      }

      if (fontSize[size]) {
        fontSize[size]++;
      } else {
        fontSize[size] = 1;
      }
    }
    
  });

  // Find the font with the maximum count
  var mostUsedFont = Object.keys(fontCounts).reduce(function(a, b) {
    return fontCounts[a] > fontCounts[b] ? a : b;
  });

  var mostUsedSize = Object.keys(fontSize).reduce(function(a, b) {
    return fontSize[a] > fontSize[b] ? a : b;
  });

  const style = {}
  style[DocumentApp.Attribute.FONT_SIZE] = mostUsedSize ;
  style[DocumentApp.Attribute.FONT_FAMILY] = mostUsedFont;
  style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
  style[DocumentApp.Attribute.FOREGROUND_COLOR] = '#000000';

  // Return the most used font
  return style;
}

function setTimes() {
  fontFam = 'Times New Roman'
}

function isFontFamilyMatch(input) {
  var fontFamilies = DocumentApp.Attribute.FONT_FAMILY;
  
  for (var fontFamily in fontFamilies) {
    if (fontFamilies[fontFamily] === input) {
      return true;
    }
  }
  
  return false;
}
