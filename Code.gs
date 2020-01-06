var CHECKED_LABEL_NAME = "Spellchecked";
var ISSUE_LABEL_NAME = "Spelling Issue";

function main() {
  Logger.log('debug: creating bing spell checker object');
  var bing = new BingSpellChecker({
    key : 'xxxxxxxxxxxxxxxx',
    toIgnore : ['adwords','adgroup','russ'],
    enableCache : true
  });
  Logger.log('debug: getting mccapp acounts');
  var accountIter = MccApp.accounts().get();
  while(accountIter.hasNext()) {
    Logger.log('debug: selecting account');
    MccApp.select(accountIter.next());
    Logger.log('debug: checking ads');
    checkAds(bing);
    if(bing.hitQuota) {
      Logger.log('debug: bing quota hit');
      break;
    }
  }
  Logger.log('debug: saving cache');
  bing.saveCache();
}

function checkAds(bing) {
  createLabelIfNeeded(CHECKED_LABEL_NAME,'Indicates an entity was spell checked','#00ff00' /*green*/);
  createLabelIfNeeded(ISSUE_LABEL_NAME,'Indicates an entity has a spelling issue','#ff0000' /*red*/);

  Logger.log('debug: creating ad iter');
  var adIter = AdWordsApp.ads()
      .withCondition("Status = ENABLED")
      .withCondition(Utilities.formatString("LabelNames CONTAINS_NONE ['%s','%s']",
          CHECKED_LABEL_NAME,
          ISSUE_LABEL_NAME))
      .get();
  while(adIter.hasNext() && !bing.hitQuota) {
    var ad = adIter.next();
    var textToCheck = "";
    if (ad.getType() === "EXPANDED_TEXT_AD") {
      var expandedTextAd = ad.asType().expandedTextAd();
      textToCheck = [
        expandedTextAd.getHeadlinePart1(),
        expandedTextAd.getHeadlinePart2(),
        expandedTextAd.getDescription()
      ].join(' ');
    } else {
      textToCheck = [
        ad.getHeadline(),
        ad.getDescription1(),
        ad.getDescription2()
      ].join(' ');
    }
    try {
      Logger.log('debug: checking text: '+textToCheck);
      var hasSpellingIssues = bing.hasSpellingIssues(textToCheck);
      if(hasSpellingIssues) {
        ad.applyLabel(ISSUE_LABEL_NAME);
      } else {
        ad.applyLabel(CHECKED_LABEL_NAME);
      }
    } catch(e) {
      // This probably means you're out of quota.
      // You can pick up from here next time.
      Logger.log('INFO: '+e);
      break;
    }
    if(!AdWordsApp.getExecutionInfo().isPreview() &&
        AdWordsApp.getExecutionInfo().getRemainingTime() < 60) {
      // Out of time
      Logger.log("INFO: Ran out of time. Will continue next run.");
      break;
    }
  }
}

//This is a helper function to create the label if it does not already exist
function createLabelIfNeeded(name,description,color) {
  if(!AdWordsApp.labels().withCondition("Name = '"+name+"'").get().hasNext()) {
    Logger.log('debug: creating label: ' + name);
    AdWordsApp.createLabel(name,description,color);
  }
}

/******************************************
 * Bing Spellchecker API v1.0
 * By: Russ Savage (@russellsavage)
 * Usage:
 *  // You will need a key from
 *  // https://www.microsoft.com/cognitive-services/en-us/bing-spell-check-api/documentation
 *  // to use this library.
 *  var bing = new BingSpellChecker({
 *    key : 'xxxxxxxxxxxxxxxxxxxxxxxxx',
 *    toIgnore : ['list','of','words','to','ignore'],
 *    enableCache : true // <- stores data in a file to reduce api calls
 *  });
 * // Example usage:
 * var hasSpellingIssues = bing.hasSpellingIssues('this is a speling error');
 ******************************************/
function BingSpellChecker(config) {
  this.BASE_URL = 'https://api.cognitive.microsoft.com/bing/v7.0/spellcheck';
  this.CACHE_FILE_NAME = 'spellcheck_cache.json';
  this.key = config.key;
  this.toIgnore = config.toIgnore;
  this.cache = null;
  this.previousText = null;
  this.previousResult = null;
  this.delay = (config.delay) ? config.delay : 60000/7;
  this.timeOfLastCall = null;
  this.hitQuota = false;

  Logger.log('debug: BingSpellChecker initialization');

  // Given a set of options, this function calls the API to check the spelling
  // options:
  //   options.text : the text to check
  //   options.mode : the mode to use, defaults to 'proof'
  // returns a list of misspelled words, or empty list if everything is good.
  this.checkSpelling = function(options) {
    if(this.toIgnore) {
      options.text = options.text.replace(new RegExp(this.toIgnore.join('|'),'gi'), '');
    }
    options.text = options.text.replace(/{.+}/gi, '');
    options.text = options.text.replace(/[^a-z ]/gi, '').trim();

    if(options.text.trim()) {
      if(options.text === this.previousText) {
        Logger.log('INFO: Using previous response.');
        return this.previousResult;
      }
      if(this.cache) {
        var words = options.text.split(/ +/);
        for(var i in words) {
          Logger.log('INFO: checking cache: '+words[i]);
          if(this.cache.incorrect[words[i]]) {
            Logger.log('INFO: Using cached response.');
            return [{"offset":1,"token":words[i],"type":"cacheHit","suggestions":[]}];
          }
        }
      }
      var url = this.BASE_URL;
      var config = {
        method : 'POST',
        headers : {
          'Ocp-Apim-Subscription-Key' : this.key,
          'Content-Type' : 'application/x-www-form-urlencoded'
        },
        payload : 'Text='+encodeURIComponent(options.text),
        muteHttpExceptions : true
      };
      if(options && options.mode) {
        url += '?mode='+options.mode;
      } else {
        url += '?mode=proof';
      }
      if(this.timeOfLastCall) {
        var now = Date.now();
        if(now - this.timeOfLastCall < this.delay) {
          Logger.log(Utilities.formatString('INFO: Sleeping for %s milliseconds',
              this.delay - (now - this.timeOfLastCall)));
          Utilities.sleep(this.delay - (now - this.timeOfLastCall));
        }
      }
      Logger.log('debug: fetching url');
      Logger.log(url);
      Logger.log(config);
      var resp = UrlFetchApp.fetch(url, config);
      Logger.log(resp.getResponseCode());
      Logger.log(resp.getAllHeaders());
      Logger.log(resp);
      this.timeOfLastCall = Date.now();
      if(resp.getResponseCode() !== 200) {
        if(resp.getResponseCode() === 403) {
          this.hitQuota = true;
        }
        throw JSON.parse(resp.getContentText()).message;
      } else {
        var jsonResp = JSON.parse(resp.getContentText());
        this.previousText = options.text;
        this.previousResult = jsonResp.flaggedTokens;
        for(var i in jsonResp.flaggedTokens) {
          this.cache.incorrect[jsonResp.flaggedTokens[i].token] = true;
        }
        return jsonResp.flaggedTokens;
      }
    } else {
      return [];
    }
  };

  // Returns true if there are spelling mistakes in the text toCheck
  // toCheck : the phrase to spellcheck
  // returns true if there are words misspelled, false otherwise.
  this.hasSpellingIssues = function(toCheck) {
    var issues = this.checkSpelling({ text : toCheck });
    Logger.log('issues found: ' +issues);
    return (issues.length > 0);
  };

  // Loads the list of misspelled words from Google Drive.
  // set config.enableCache to true to enable.
  this.loadCache = function() {
    Logger.log('debug: loading cache');
    var fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      this.cache = JSON.parse(fileIter.next().getBlob().getDataAsString());
      Logger.log('debug: this.cache:');
      Logger.log(this.cache);
    } else {
      this.cache = { incorrect : {} };
    }
  };

  if(config.enableCache) {
    this.loadCache();
  }

  // Called when you are finished with everything to store the data back to Google Drive
  this.saveCache = function() {
    var fileIter = DriveApp.getFilesByName(this.CACHE_FILE_NAME);
    if(fileIter.hasNext()) {
      fileIter.next().setContent(JSON.stringify(this.cache));
    } else {
      DriveApp.createFile(this.CACHE_FILE_NAME, JSON.stringify(this.cache));
    }
  }
}