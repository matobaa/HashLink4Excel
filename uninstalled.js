var WS_NOTVISIVLE   = 0;
var WS_ACT_NOMAL    = 1; // default
var WS_ACT_MIN      = 2;
var WS_ACT_MAX      = 3;
var WS_NOTACT_NOMAL = 4;
var WS_ACT_DEF      = 5;
var WS_NOTACT_MIN   = 7;
var sh = new ActiveXObject( "WScript.Shell" );
sh.Run( "http://matobaa.github.io/questionnaire.html?prod=hashlink4excel&ver=0.5", WS_ACT_MIN );
sh = null;