var XLSX = require('xlsx');
var fs = require('fs');
var querystring = require('querystring');
var async = require('async');
var Bitly = require('bitly');

if (process.argv.length == 4) {
  processTweets(process.argv[2],process.argv[3],null,exitApp);
} else if (process.argv.length == 5) {
  processTweets(process.argv[2],process.argv[3],process.argv[4],exitApp);
} else {
  usage();
}

function exitApp(err) {
  if (err) {
    console.error(err);
  }
  process.exit(err == null ? 0 : -1);
}

function usage() {
  console.log('Must pass in source XLSX file, destination XLSX filename, and an optional Bit.ly generic access token.');
}

function processTweets(inFile,outFile,bitlyToken,done) {
  async.waterfall([
    function(next) {
      fs.readFile(inFile,function(err,data) {
        next(err,data);
      });
    },
    function(data,next) {
      // try {
        var workbook = XLSX.read(data);
        var firstSheet = workbook.SheetNames[0];
        var worksheet = workbook.Sheets[firstSheet];
        var _nameCol = 'A';
        var _tweetCol = 'B';
        var _startRow = 2;
        var row = _startRow;
        var tweetsAndNames = [];
        while(worksheet[_nameCol+row]) {
          tweetsAndNames.push({
            'name': worksheet[_nameCol+row].v,
            'tweet': worksheet[_tweetCol+row].v
          })
          row++;
        }
        next(null,tweetsAndNames);
      // } catch(e) {
      //   next(e);
      // }
    },
    function(tweetsAndNames,next) {
      var bitly;
      if (bitlyToken) {
        bitly = new Bitly(bitlyToken);
      }
      async.parallel(
        tweetsAndNames.map(function(tweetAndName) {
          return function(next1) {
            var url = 'http://twitter.com/home?' + querystring.stringify({
              'status': tweetAndName.tweet
            });
            if (bitly) {
              bitly.shorten(url).then(function(response) {
                tweetAndName.link = response.data.url;
                next1(null,tweetAndName);
              },next1);
            } else {
              tweetAndName.link = url;
              next1(null,tweetAndName);
            }
          }
        }),
        next
      );
    },
    function(tweetsAndLinksAndNames,next) {
      try {
        var wb = {
          'SheetNames': ['Tweets'],
          'Sheets': {
            'Tweets': {
              '!ref': 'A1:C' + (tweetsAndLinksAndNames.length + 1),
              '!merges': [],
              'A1': {
                'v': 'Name',
                't': 's'
              },
              'B1': {
                'v': 'Tweet',
                't': 's'
              },
              'C1': {
                'v': 'Link',
                't': 's'
              }
            }
          }
        };
        tweetsAndLinksAndNames.forEach(function(tweetAndLinkAndName,index) {
          wb.Sheets.Tweets['A' + (index+2)] = {
            'v': tweetAndLinkAndName.name,
            't': 's'
          };
          wb.Sheets.Tweets['B' + (index+2)] = {
            'v': tweetAndLinkAndName.tweet,
            't': 's'
          };
          wb.Sheets.Tweets['C' + (index+2)] = {
            'v': tweetAndLinkAndName.link,
            't': 's'
          };
        });
        var data = XLSX.write(wb,{
          'bookType': 'xlsx',
          'bookSST': false,
          'type': 'buffer'
        });
        fs.writeFile(outFile,data,function(err) {
          next(err);
        });
      } catch(e) {
        next(e);
      }
    }
  ],done);
}
