const csv = require("csv");
const Excel = require("exceljs");

class CSV {
    /**
     *
     * @param {Array} data - Array of objects or arrays.
     * @param {Array} columns - Column headers.
     */
    constructor(data, columns) {
        this.data = data;
        this.columns = columns;
    }

    /**
     * @callback writeCallback
     * @param {Object} err - Error output.
     * @param {String} [fileName] - Name of saved file.
     * @param {String} [dir] - Directory of saved file.
     */

    /**
     * Converts data object into
     * @param {String} dir - Directory to save CSV file.
     * @param {String} fileName - Name of CSV file to save.
     * @param {writeCallback} callback.
     */
    write(dir, fileName, callback) {
        let d = Format.date();
        let fileDate = d.year + d.month + d.date + "_" + d.hours + d.minutes + d.seconds;
        let options = this.columns ? {header: true, columns: this.columns} : null;
        let csvData = this.data;

        if (options) {
            csvData = objectToArray(csvData, this.columns);
        }

        csv.stringify(this.data, options, function(err, data) {
            if (err) {
                callback({message: "Unable to stringify CSV file.", err: err});
            } else {
                FileSystem.read(dir, function(err) {
                    if (err) {
                        callback(err);
                    } else {
                        FileSystem.writeFile(dir + fileName + "_" + fileDate + ".csv", data, function(err) {   //write CSV file
                            if (err) {
                                callback(err);
                            } else {
                                callback(null, fileName + "_" + fileDate + ".csv", dir);
                            }
                        });
                    }
                });
            }
        });
    }

    /**
     * @callback readCallback
     * @param {Object} err - Error output.
     * @param {*} [items] - Object or array of data in CSV file.
     */

    /**
     * Reads a directory and returns an array of file names.
     * @param {String|Array} input - The path or data of the file to read.
     * @param {readCallback} callback.
     * @param {Object|Function} [options] - Various file options.
     * @param {Boolean} [options.headers] - If the file contains headers in the first row.
     * @param {Boolean} [options.lock] - If the file should be locked before read.
     */
    static read(input, callback, options) {
        switch (true) {
            case Array.isArray(path):
                parseCSV(input, callback, options);
                break;

            case FileSystem.isDir(input):
                readDirectory(input, callback, options);
                break;

            default:
                readFile(input, callback, options);
                break;
        }

        function readDirectory(dir, callback, options) {
            FileSystem.read(dir, ["csv", "xlsx", "xls"], function (err, files) {
                if (err) {
                    callback(err);
                } else if (files.length > 0) {
                    let files = [];

                    function readFiles(i) {
                        readFile(dir + files[i], function (err, data, columnHeaders) {
                            if (err) {
                                callback(err);
                            } else {
                                files.push({file: dir + files[i], data: data, columnHeaders: columnHeaders});

                                if (i+1 < files.length) {
                                    readFiles(i+1);
                                } else {
                                    callback(null, files);
                                }
                            }
                        }, options);
                    }

                    readFiles(0);    //start with first file
                } else {
                    callback(null, []);
                }
            });
        }

        function readFile(path, callback, options) {
            lockFile(path, options, function(err, path) {
                if (err) {
                    throw new Err(err);
                } else {
                    switch (getFileExt(path)) {
                        case "csv":
                            FileSystem.readFile(path, function (err, file) {
                                if (err) {
                                    callback(err);
                                } else {
                                    parseCSV(file, options, callback);
                                }
                            });
                            break;

                        case "xls":
                        case "xlsx":
                            let workbook = new Excel.Workbook();
                            workbook.xlsx.readFile(path)
                                .then(function() {
                                    parseExcel(workbook, options, callback);
                                });
                            break;
                    }
                }
            })
        }

        function getFileExt(path) {
            let ext = FileSystem.getExt(path);
            return ext.toLowerCase();
        }

        function lockFile(path, options, callback) {
            if (options && options.lock) {
                FileSystem.lock(path, function (err, file, oldFile, dir) {
                    if (err) {
                        callback(err);
                    } else {
                        callback(dir + file);
                    }
                });
            } else {
                callback(path);
            }
        }
    }
}

module.exports = CSV;

function objectToArray(data, columns) {
    let newArray = [];
    for (let d=0; d<data.length; d++) {
        let item = [];

        if (!Array.isArray(data[d])) {
            for (let c=0; c<columns.length; c++) {
                item.push(data[d][columns[c]] ? data[d][columns[c]] : "");
            }
        } else {
            item = data[d];
        }

        newArray.push(item);
    }

    return newArray;
}

function parseCSV(file, callback, options) {
    csv.parse(file, function(err, rows){  //parse CSV data
        if (err) {
            callback({message: "Unable to parse \"" + file + "\".", err: err});
        } else {
            if(options.headers) {
                let items = [];
                let columnHeaders = [];

                for (let r=0; r<rows.length; r++) {
                    if (r === 0) {
                        for (let c=0; c<rows[0].length; c++) {
                            columnHeaders.push(rows[r][c]);
                        }
                    } else {
                        let item = {};
                        for (let c=0; c<rows[0].length; c++) {
                            item[rows[0][c]] = rows[r][c];
                        }
                        items.push(item);
                    }
                }

                callback(null, items, columnHeaders);
            } else {
                callback(null, rows);
            }
        }
    });
}

function parseExcel(file, callback, options) {
    let rows = file.worksheets[0]._rows;

    let items = [];
    let item;
    let columnHeaders = [];

    for (let r=0; r<rows.length; r++) {
        if (rows[r].hasValues) {
            let columns = rows[r]._cells;
            if (options && options.headers) {
                item = {};
            } else {
                item = [];
            }
            for (let c=0; c<columns.length; c++) {
                if (columns[c]) {   //Sometimes the extension misses out indexes where columns have been deleted
                    if (options && options.headers) {
                        if (r === 0) {
                            columnHeaders.push(columns[c].value);
                        } else {
                            if (columns[c] && (columns[c].value === 0 || columns[c].value === "0" || columns[c].value)) {
                                if (typeof columns[c].value === "string") {
                                    columns[c].value = columns[c].value.trim();
                                }
                                item[columnHeaders[c]] = columns[c].value;
                            } else {
                                item[columnHeaders[c]] = "";
                            }
                        }
                    } else {
                        if (columns[c] && (columns[c].value === 0 || columns[c].value === "0" || columns[c].value)) {
                            if (typeof columns[c].value === "string") {
                                columns[c].value = columns[c].value.trim();
                            }
                            item.push(columns[c].value);
                        } else {
                            item.push("");
                        }
                    }
                }
            }

            if (options && options.headers && r === 0) {
                //do nothing
            } else {
                items.push(item);
            }
        }
    }
    callback(null, items, columnHeaders);
}
