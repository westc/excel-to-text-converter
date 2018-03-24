const {shell, remote} = require('electron');
const {dialog, app} = remote;
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

const DEFAULT_NAMING_SCHEME = '<WORKBOOK> - <SHEET>.csv';
const RGX_META = /<(WORKBOOK|SHEET|SHEET_NUMBER|EXTENSION)>/g;

function recurseDirSync(currentDirPath, opt_filter) {
  let result = {
      isFile: false,
      path: currentDirPath,
      stat: fs.statSync(currentDirPath),
      files: []
    };

  fs.readdirSync(currentDirPath).forEach(function (name) {
    let filePath = path.join(currentDirPath, name),
        stat = fs.statSync(filePath),
        isFile = stat.isFile();
    if ((isFile || stat.isDirectory()) && (!opt_filter || opt_filter(filePath, isFile, stat))) {
      result.files.push(isFile ? { isFile: true, path: filePath, stat: stat } : recurseDirSync(filePath, opt_filter));
    }
  });
  return result;
}

function ensureDirExists(dirPath) {
  if (!fs.existsSync(dirPath)) {
    ensureDirExists(path.dirname(dirPath));
    fs.mkdirSync(dirPath);
  }
}

window.recurseDirSync = recurseDirSync;

$(function() {
  $('#app-title').text(document.title);

  let myVue = new Vue({
    el: '#myVue',
    data: {
      folderPath: null,
      workbooks: [],
      sheetFilters: [{ value: '', isRegExp: false, error: null }],
      converting: false,
      namingScheme: DEFAULT_NAMING_SCHEME,
      DEFAULT_NAMING_SCHEME: DEFAULT_NAMING_SCHEME
    },
    watch: {
      sheetFilters: {
        deep: true,
        handler() {
          var i = this.sheetFilters.findIndex(f => !f.value);
          if (this.sheetFilters.findIndex(f => !f.value) < 0) {
            this.sheetFilters.push({ value: '', isRegExp: false, error: null });
          }
        }
      }
    },
    computed: {
      includedWorkbooks() {
        return this.workbooks.filter(function(wb) {
          return wb.include;
        });
      },
      normalizedNamingScheme() {
        return this.namingScheme.replace(/[\\\/]+/g, path.sep);
      },
      namingSchemeError() {
        let namingScheme = this.normalizedNamingScheme;
        return namingScheme.endsWith(path.sep)
          ? 'Cannot end with "' + path.sep + '".'
          : /(^|[\\\/])\.\./.test(namingScheme)
            ? 'Cannot start with "..".'
            : /[~#%&*\{\}:\?\+\|"]/.test(namingScheme)
              ? 'Cannot contain any of the following:  ~ # & * { } : ? + |'
              : /[<>]/.test(namingScheme.replace(RGX_META, ''))
                ? 'Cannot contain the less than (<) or greater than (>) characters except for when adding "<WORKBOOK>", "<SHEET>", and "<SHEET_NUMBER>".'
                : !RGX_META.test(namingScheme)
                  ? 'Must use at least one of the following:  <WORKBOOK>, <SHEET>, <SHEET_NUMBER>, <EXTENSION>'
                  : false;
      },
      isValidNamingScheme() {
        return !this.namingSchemeError;
      },
      hasInvalidFilters() {
        return this.sheetFilters.filter(f => f.error).length > 0;
      },
      canConvert() {
        let myVue = this;
        return myVue.workbooks.length
          && myVue.normalizedNamingScheme
          && !myVue.namingSchemeError
          && !myVue.hasInvalidFilters
          && myVue.preConversions.filter(file => myVue.matchesFilters(file.sheetName)).length;
      },
      workbookSheets() {
        let result = [],
            myVue = this;
        myVue.includedWorkbooks.forEach(objWB => {
          let wb = XLSX.readFileSync(objWB.path),
              sheetNumber = 0;
          wb.Workbook.Sheets.forEach(sheetDetails => {
            if (sheetDetails.Hidden === 0) {
              result.push({
                workbook: wb,
                workbookPath: objWB.path,
                sheetName: sheetDetails.name,
                sheet: wb.Sheets[sheetDetails.name],
                sheetDetails,
                sheetNumber: ++sheetNumber
              });
            }
          });
        });
        return result;
      },
      preConversions() {
        let myVue = this,
            ext = path.extname(myVue.normalizedNamingScheme),
            typeName = /\.json$/i.test(ext)
            ? 'json'
            : /\.(txt|tsv)$/i.test(ext)
              ? 'tsv'
              : 'csv';
        return this.workbookSheets.map(o => JS.extend(o, {
          newFilePath: myVue.getSheetFilePath(o.workbookPath, o.sheetName, o.sheetNumber),
          typeName
        }));
      }
    },
    methods: {
      matchesFilters(sheetName) {
        let myVue = this,
            filters = myVue.sheetFilters.filter(f => f.isRegExp ? JS.isRegExp(myVue.parseRegExp(f.value)) : f.value),
            i = filters.length;
        if (!i) {
          return true;
        }
        for (; i--; ) {
          let filter = filters[i];
          if (filter.isRegExp ? myVue.parseRegExp(filter.value).test(sheetName) : (filter.value == sheetName)) {
            return true;
          }
        }
        return false;
      },
      onChangeSheetFilter(sheetFilter) {
        var emptyIndices = [],
            sheetFilters = this.sheetFilters;
        for (let i = sheetFilters.length; i--; ) {
          if (!sheetFilters[i].value) {
            emptyIndices.push(i);
          }
        }
        if (emptyIndices.length > 1) {
          emptyIndices.slice(1).forEach(i => sheetFilters.splice(i, 1));
        }

        sheetFilter.error = null;
        if (sheetFilter.value && sheetFilter.isRegExp) {
          let rgx = this.parseRegExp(sheetFilter.value);
          if (!JS.isRegExp(rgx)) {
            sheetFilter.error = rgx;
          }
        }
      },
      parseRegExp(strRegExp) {
        try {
          let rgx;
          strRegExp = strRegExp.replace(
            /^([\|\/%\$:])((?:\\.|[^\1])+)\1(\w*)$/,
            function(m, delim, body, flags) {
              rgx = new RegExp(body, flags);
              return '';
            }
          );
          if (strRegExp) {
            throw new Error('Regular expression not recognized.  Perhaps the leading and ending delimiter was not specified.  Delimiters can be any of the following:  / | % $ :');
          }
          return rgx;
        }
        catch (e) {
          return e.message;
        }
      },
      browse() {
        dialog.showOpenDialog({properties: ['openDirectory']}, function(filePaths) {
          let dirPath = filePaths && filePaths[0];
          if (dirPath) {
            myVue.folderPath = dirPath;
            myVue.workbooks.splice(0, Infinity);

            let foundWorkbookInSub = false;
            let maxLevelsDown = 0;

            recurseDirSync(dirPath, function(filePath, isFile, stat) {
              if (/\.xlsx?$/i.test(filePath)) {
                let fileDirPath = path.dirname(filePath);
                if (!foundWorkbookInSub && (fileDirPath != dirPath)) {
                  foundWorkbookInSub = true;
                  maxLevelsDown = dialog.showMessageBox({
                    title: 'How Many Levels Down',
                    message: 'How many levels down would you like to search for Excel files?',
                    buttons:['Selected Only', '1', '2', 'All'],
                    defaultId: 0,
                    type: 'question'
                  });
                  maxLevelsDown = maxLevelsDown < 3 ? maxLevelsDown : Infinity;
                }

                let levelsDown = fileDirPath.slice(dirPath.length).split(path.sep).length - 1;
                if (levelsDown <= maxLevelsDown) {
                  myVue.workbooks.push({
                    path: filePath,
                    include: true
                  });
                }
              }
              return true;
            });
          }
        });
      },
      getSheetFilePath(filePath, sheetName, sheetNumber) {
        let extName = path.extname(filePath);
        let wbName = path.basename(filePath).slice(0, -extName.length);
        extName = extName.replace(/^\./, '');

        let newPath = /^[\\\/]/.test(this.normalizedNamingScheme)
          ? this.folderPath + this.normalizedNamingScheme
          : (path.dirname(filePath) + path.sep + this.normalizedNamingScheme);
        return newPath.replace(
          RGX_META,
          function(m, typeName) {
            return typeName == 'WORKBOOK'
              ? wbName
              : typeName == 'SHEET'
                ? sheetName
                : typeName == 'SHEET_NUMBER'
                  ? sheetNumber
                  : extName;
          }
        );
      },
      convert() {
        let myVue = this;
        myVue.preConversions.forEach(file => {
          myVue.converting = true;
          if (myVue.matchesFilters(file.sheetName)) {
            let output = XLSX.utils['json' == file.typeName ? 'sheet_to_json' : 'sheet_to_csv'](
              file.sheet,
              { FS: 'tsv' == file.typeName ? '\t' : ',' }
            );

            ensureDirExists(path.dirname(file.newFilePath));

            fs.writeFileSync(
              file.newFilePath,
              'string' == typeof output ? output : JSON.stringify(output)
            );
          }
          myVue.converting = false;
        });

        $('#modalMessage').modal('show');
      },
      visitBlog() {
        shell.openExternal("http://www.cwestblog.com");
      },
      close() {
        app.exit();
      }
    },
    updated() {
      $('[title]').tooltip();
    },
    mounted() {
      $('[title]').tooltip();
    }
  });
});
