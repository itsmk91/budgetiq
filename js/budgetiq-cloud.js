(function () {
  'use strict';
  if (window.location.hostname !== 'itsmk91.github.io') return;
  const production = 'https://budgetiq-iphone-test.mc2qww2kxf.chatgpt.site';
  let path = window.location.pathname.replace(/^\/budgetiq/, '') || '/splash';
  if (path === '/') path = '/splash';
  path = path.replace(/\.html$/, '');
  window.location.replace(`${production}${path}${window.location.search}${window.location.hash}`);
})();
