// Tiny unit test framework for JScript
var test = function (component, fn, count) {
  count = count || 1;
  WScript.Echo('# ' + component);

  fn({
    same: function (actual, expected, msg) {
      if (actual == expected) {
        WScript.Echo('pass ' + count + ' - ' + msg);
      } else {
        throw new Error(
          '\nfail ' + count + ' - ' + msg + '\n expected:\n  ' + expected + '\n actual:\n  ' + actual
          );
      }
      count++;
    }
  });
};
