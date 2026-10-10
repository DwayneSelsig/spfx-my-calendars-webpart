const spfxReactProfile = require('@microsoft/eslint-config-spfx/lib/flat-profiles/react');
const projectConfig = require('./eslint-project-rules.js');
const reactHooksPlugin = require('eslint-plugin-react-hooks');

module.exports = [
  ...spfxReactProfile.map(config => config.plugins && config.plugins['react-hooks']
    ? { ...config, plugins: { ...config.plugins, 'react-hooks': reactHooksPlugin } }
    : config),
  {
    // Preserve the legacy config's handling of existing disable directives.
    linterOptions: { reportUnusedDisableDirectives: false },
    files: ['**/*.ts', '**/*.tsx'],
    // The legacy SPFx profile disables this rule; project rules use no-var-requires.
    rules: { '@typescript-eslint/no-require-imports': 'off' }
  },
  ...projectConfig.overrides.map(override => ({
    // Legacy basename patterns matched files in every directory.
    files: override.files.map(pattern => pattern.includes('/') ? pattern : `**/${pattern}`),
    ...(override.parser ? {
      languageOptions: {
        parser: require(override.parser),
        parserOptions: {
          ...projectConfig.parserOptions,
          ...override.parserOptions
        }
      }
    } : {}),
    rules: override.rules
  }))
];
