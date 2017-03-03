// Tiny unit test framework for JScript
var test = function (component, fn, count) {

  count = count || 1;
  console.log('\n# ' + component);

  fn({
    same: function (actual, expected, msg) {
      if (actual == expected) {
        console.log('pass ' + count + ' - ' + msg);
      } else {
        throw new Error(
          '\nfail ' + count + ' - ' + msg +
          '\n expected:' +
          '\n  ' + expected +
          '\n actual:' +
          '\n  ' + actual
          );
      }
      count++;
    }
  });
};
