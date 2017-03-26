#!/usr/bin/env node

// node changePassword.js exchange-host(如mail.baidu.com) 用户名 密码
// node changePassword.js mail.corp.qunar.com wenjie.zhang Qunar.987

/*
 * Created with WebStorm.
 * @author: soncy
 * @date: 12/16/13 12:12 PM
 * @contact: soncy1986@gmail.com
 * @fileoverview: 更改microsoft exchange密码
 */

var url = require('url');
var https = require('https');
var querystring = require('querystring');
var co = require('co');
var prompt = require('co-prompt');
var program = require('commander');
var chalk = require('chalk');
var ProgressBar = require('progress');
var cookie = require('./lib/cookieManager');
var request = require('./lib/requestManager');

function range(val) {
  return val.split('..').map(Number);
}

function list(val) {
  return val.split(',');
}

function collect(val, memo) {
  memo.push(val);
  return memo;
}

function increaseVerbosity(v, total) {
  return total + 1;
}


program
  .version('1.0.0')
  .option('-H, --host [host]', '保持密码不变的站点 host', 'mail.corp.qunar.com')
  // .option('-p, --password <password>', 'The user\'s password')
  // .action(function(file) {
  //   console.log('user: %s pass: %s file: %s', program.username, program.password, file);
  // })
  .parse(process.argv);

var MAXCOUNT = 5;
var passwordTemp = [];
var h = program.host;
var host = "https://" + h;
var userName = '';
var targetPassword = '';
var charList = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"; // 字母表，用于生成随机密码
var specialChar = '!@#$';
var currentPassword = targetPassword;
var changeCount = 0;
var barOpts = {
      width: 20,
      total: 120,
      clear: true
    };
var bar = new ProgressBar(' doing [:bar] :percent :etas', barOpts);

(function keepPassword() {
    co(function *() {
      userName = yield prompt('username: ');
      targetPassword = yield prompt.password('password: ');
      currentPassword = targetPassword;
      // console.log('user: %s pass: %s', userName, targetPassword);
      login();
    });
})()

function createRandomChar() {
  var randomNumber = parseInt(Math.random() * 1000000);
  var randomString = charList[parseInt(Math.random() * 51)] + charList[parseInt(Math.random() * 51)];
  var randomSpecial = specialChar[parseInt(Math.random() * 3)];
  return randomString + randomSpecial + randomNumber;
}

function getNewPassword() {
  var password = createRandomChar();
  if (~passwordTemp.indexOf(password)) {
    newPassword();
  }
  // console.log('生成新密码：' + password);
  return password;
}

function login() {
  var params = {
    username: userName,
    password: currentPassword,
    trusted: 0,
    isUtf8: 1,
    forcedownlevel: 0,
    flags: 0,
    destination: host + '/owa/'
  };

  var headers = {
    "Referer": host + '/owa/auth/logon.aspx?replaceCurrent=1&url=https%3a%2f%2f' + h + '%2fowa%2f',
    "Host": url.parse(host).host,
    "Cookie": "PBack=0",
    "Content-Type": "application/x-www-form-urlencoded",
    "Content-length": querystring.stringify(params).length
  };
  var urlString = host + '/owa/auth.owa';
  request.go('POST', urlString, querystring.stringify(params), headers, function (res) {
    cookie.set(res.headers['set-cookie']);
    visitChangePage();
  });
}

function visitChangePage() {

  var headers = {
    "Cookie": cookie.getAll(),
    "Referer": host + '/owa/?ae=Options&t=ChangePassword'
  };

  var urlString = host + '/owa/?ae=Options&t=ChangePassword';
  request.go('GET', urlString, '', headers, function (res) {
    cookie.set(res.headers['set-cookie']);
    changePassword();
  });
}

function changePassword() {
  var userContext = cookie.get('UserContext');
  var oldPassword = currentPassword;
  var newPassword = changeCount === MAXCOUNT ? targetPassword : getNewPassword();
  var params = '<params><canary>' + userContext + '</canary><oldPwd>'
    + oldPassword + '</oldPwd><newPwd>' + newPassword + '</newPwd></params>';

  var headers = {
    "Cookie": cookie.getAll(),
    "Referer": host + '/owa/?ae=Options&t=ChangePassword',
    "Content-Type": "text/plain; charset=UTF-8",
    "Content-Length": params.length
  };

  var urlString = host + '/owa/ev.owa?oeh=1&ns=Options&ev=ChangePassword';
  // console.log('密码将被改为：' + newPassword);
  request.go('POST', urlString, params, headers, function (res, content) {
    if (content === '') {
      // console.log('密码已经被改为：' + newPassword);
      changeCount++;
      bar.tick(20);
      currentPassword = newPassword;

      if (changeCount > MAXCOUNT) {
        // console.log('密码已经被改回：' + targetPassword);
        console.log(chalk.bold.cyan('密码延期成功'));
        process.exit();
      }
      cookie.clear();
      setTimeout(login, 1000);
    } else {
      console.log(content);
    }
  });
}

process.on('exit', function() {
  // console.log(123);
});

process.on('SIGINT', function() {
  // console.log("Ignored Ctrl-C");
  process.exit(0);  
});
// login();
