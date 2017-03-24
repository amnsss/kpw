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

var cookie = require('./lib/cookieManager');
var request = require('./lib/requestManager');
var program = require('commander');

program
 .arguments('<file>')
 .option('-u, --username <username>', 'The user to authenticate as')
 .option('-p, --password <password>', 'The user\'s password')
 .action(function(file) {
   console.log('user: %s pass: %s file: %s',
       program.username, program.password, file);
 })
 .parse(process.argv);

var MAXCOUNT = 5;
var passwordTemp = [];
var h = process.argv[2];
var host = "https://" + h;
var userName = process.argv[3];
var targetPassword = process.argv[4];

var charList = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"; // 字母表，用于生成随机密码
var specialChar = '!@#$';
var currentPassword = targetPassword;
var changeCount = 0;

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
  console.log('生成新密码：' + password);
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
  console.log('密码将被改为：' + newPassword);
  request.go('POST', urlString, params, headers, function (res, content) {
    if (content === '') {
      console.log('密码已经被改为：' + newPassword);
      changeCount++;
      currentPassword = newPassword;

      if (changeCount > MAXCOUNT) {
        console.log('密码已经被改回：' + targetPassword);
        process.exit();
      }
      cookie.clear();
      setTimeout(login, 1000);
    } else {
      console.log(content);
    }
  });
}
// login();
