// Something to test
var double = function (x) { return x * 2; }

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

test('double', function (assert) {
  {
    var msg = 'double() should take a number x and return the product of x and 2';

    var actual = double(4);
    var expected = 8;

    assert.same(actual, expected, msg);
  }

  {
    var msg = 'double() should return NaN for non-numbers';

    var actual = isNaN(double('puppy'));
    var expected = true;

    assert.same(actual, expected, msg);
  }

  // failing test
  {
    var msg = 'false should be true?';

    var actual = false;
    var expected = true;

    assert.same(actual, expected, msg);
  }
});