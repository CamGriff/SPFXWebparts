module.exports = function (config) {
  // Remove any existing rule already targeting .html files (the base
  // SPFx/Heft build configuration includes a default html-loader rule),
  // so it doesn't chain together with our own rule below and fail trying
  // to parse Handlebars syntax as strict HTML.
  config.module.rules = config.module.rules.filter(function (rule) {
    return !(rule.test && rule.test.toString().indexOf('html') !== -1);
  });

  config.module.rules.push({
    test: /\.html$/,
    use: 'raw-loader'
  });

  return config;
};