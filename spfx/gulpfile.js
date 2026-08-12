'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppressRule(/Warning - \[sass\]/gi);
build.initialize(require('gulp'));
