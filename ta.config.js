// Default config — safe to commit. Override in ta.config.local.js (gitignored).
module.exports = {
  TA_API_BASE_URL: '',
  TA_SDK_BASE_URL: '',          // Optional: defaults to TA_API_BASE_URL + '/TotalAgility/Services/Sdk'
  TA_PROCESS_NAME: 'Word Doc Review',
  TA_AUTH_TYPE: 'federated',   // 'federated' | 'password' | 'windows'
  TA_API_TOKEN: '',             // Optional Bearer token for custom REST endpoints
};
