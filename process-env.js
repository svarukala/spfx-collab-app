"use strict";
 
const fs = require("fs");
const path = require("path");
const dve = require("dotenv-expand");
const de = require("dotenv")

const NODE_ENV = process.env.NODE_ENV || "dev";
const dotEnvPath = path.resolve(process.cwd(), ".env");
 
const dotenvFiles = [
  `${dotEnvPath}.${NODE_ENV}.local`,
  `${dotEnvPath}.${NODE_ENV}`,
  NODE_ENV !== "test" && `${dotEnvPath}.local`,
  dotEnvPath
].filter(Boolean);
 
dotenvFiles.forEach(dotenvFile => {
  if (fs.existsSync(dotenvFile)) {
    
      dve.expand(
      de.config({
        path: dotenvFile
      })
    );
  }
});
 
const appDirectory = fs.realpathSync(process.cwd());
process.env.NODE_PATH = (process.env.NODE_PATH || "")
  .split(path.delimiter)
  .filter(folder => folder && !path.isAbsolute(folder))
  .map(folder => path.resolve(appDirectory, folder))
  .join(path.delimiter);
 
const SPFX_ = /^SPFX_/i;
 
function getClientEnvironment() {
  const raw = Object.keys(process.env)
    .filter(key => SPFX_.test(key))
    .reduce((env, key) => {
      env[key] = process.env[key];
      return env;
    }, {});
 
  const stringified = {
    "process.env": Object.keys(raw).reduce((env, key) => {
      env[key] = JSON.stringify(raw[key]);
      return env;
    }, {})
  };
 
  return { raw, stringified };
}
 
module.exports = getClientEnvironment;