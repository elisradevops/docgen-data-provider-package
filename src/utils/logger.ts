'use strict';
import * as winston from 'winston';
import BrowserConsole from 'winston-transport-browserconsole';
let logger: winston.Logger;

const logFormat = winston.format.printf((info) => `${info.timestamp} - ${info.level}: ${info.message}`);
// if (typeof window === "undefined") {
//   let logsPath = process.env.logs_path || "./logs/";

//   logger = winston.createLogger({
//     format: winston.format.timestamp(),
//     level: "silly",
//     transports: [
//       new winston.transports.File({
//         filename: `${logsPath}azure-rest-api-errors.log`,
//         level: "error",
//         format: logFormat,
//       }),
//       new winston.transports.File({
//         filename: `${logsPath}azure-rest-api-all.log`,
//         format: logFormat,
//       }),
//       new winston.transports.Console({ format: logFormat, level: "debug" }),
//     ],
//   });
// } else {
logger = winston.createLogger({
  format: winston.format.timestamp(),
  level: 'silly',
  transports: [
    new winston.transports.Console({ format: logFormat, level: 'debug' }),
    // new BrowserConsole({ format: logFormat, level: 'debug' }),
  ],
});
// }

export default logger;
