'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

build.initialize(require('gulp'));

// update webpack config to edit constants
build.configureWebpack.mergeConfig({
  additionalConfiguration: (webpackConfig) => {
    let mergedConfig = webpackConfig;
    mergedConfig = updateBundleConstants(mergedConfig);
    return mergedConfig;
  }
});

/*
 * Replace constants through the generated bundle with those keys defined in
 * the ./config/env-[debug|prod].json files depending on the build mode.
*/
function updateBundleConstants(webpackConfig) {
  const webpack = require('webpack');

  // find define plugin
  let definePlugin;
  for (var i = 0; i < webpackConfig.plugins.length; i++) {
    if (webpackConfig.plugins[i] instanceof webpack.DefinePlugin) {
      definePlugin = webpackConfig.plugins[i];
      break;
    }
  }

  if (definePlugin) {
    // load appropriate config
    const envConfig = (webpackConfig.mode === 'development')
      ? require('./config/env-debug.json')
      : require('./config/env-prod.json');
    // load package.json & package-solution.json for additional props
    const package_properties = require('./package.json');
    const solution_properties = require('./config/package-solution.json').solution;
    // enum all properties in config & set on plugin definitions
    let keyValue = '';
    for (var key in envConfig) {
      if (envConfig[key] != '') {
        keyValue = envConfig[key];
        // if key is a placeholder to package
        if (keyValue.toLowerCase().startsWith('${{package.json|')) {
          const packageKey = keyValue.toLowerCase().split('|')[1].replace('}}', '');
          definePlugin.definitions[key] = JSON.stringify(package_properties[packageKey]);
        } else if (keyValue.toLowerCase().startsWith('${{package-solution.json|')) {
          // else if key is a placeholder to solution package
          const solutionKey = keyValue.toLowerCase().split('|')[1].replace('}}', '');
          definePlugin.definitions[key] = JSON.stringify(solution_properties[solutionKey]);
        } else {
          definePlugin.definitions[key] = JSON.stringify(keyValue);
        }
      }
    }
  }

  return webpackConfig;
}