'use strict';

if(process.argv.indexOf('dist') !== -1){
  process.argv.push('--ship');
}

const gulp = require('gulp');
const path = require('path');
const build = require('@microsoft/sp-build-web');
const bundleAnalyzer = require('webpack-bundle-analyzer');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// ********* ADDED *******
// disable tslint
build.tslintCmd.enabled = false;
// ********* ADDED *******


let syncVersionsSubtask = build.subTask('version-sync', function (gulp, buildOptions, done) {
  this.log('Synching versions');
  const gutil = require('gulp-util');
  const fs = require('fs');
  var pkgConfig = require('./package.json');
  var pkgSolution = require('./config/package-solution.json');
  this.log('package-solution.json version:\t' + pkgSolution.solution.version);
  var newVersionNumber = pkgConfig.version.split('-')[0] + '.0';
  if (pkgSolution.solution.version !== newVersionNumber) 
  {
      pkgSolution.solution.version = newVersionNumber;
      this.log('New package-solution.json version:\t' + pkgSolution.solution.version);
      fs.writeFile('./config/package-solution.json', JSON.stringify(pkgSolution, null, 4), function (err, result) 
      {
          if (err) this.log('error', err);
      });
  }
  else 
  {
      this.log('package-solution.json version is up-to-date');
  }
  done();
}); 
let syncVersionTask = build.task('version-sync', syncVersionsSubtask);
build.rig.addPreBuildTask(syncVersionTask);

build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
      const lastDirName = path.basename(__dirname);
      const dropPath = path.join(__dirname, 'temp', 'stats');
      generatedConfiguration.module.rules.push(
          {
              test: /\.(woff|woff2|ttf|eot|svg)$/,
              loader: '@microsoft/loader-cased-file',
              options: {
                  name: '[name:lower].[hash].[ext]'  
              }
          }
      );
      generatedConfiguration.plugins.push(new bundleAnalyzer.BundleAnalyzerPlugin(
        {
          openAnalyzer: false,
          analyzerMode: 'static',
          reportFilename: path.join(dropPath, `${lastDirName}.stats.html`),
          generateStatsFile: true,
          statsFilename: path.join(dropPath, `${lastDirName}.stats.json`),
          logLevel: 'error'
        }
      ));
      return generatedConfiguration;
  }
});
build.initialize(require('gulp'));

exports.dist = (done) => {
  gulp.series('bundle', 'package-solution')(done);
};
