(function () {
  var fso = new ActiveXObject("Scripting.FileSystemObject");

  test('Dependencies', function (assert) {
    (function(){
      var msg = 'JSON is an object';

      var actual = typeof JSON;
      var expected = 'object';

      assert.same(actual, expected, msg);
    }());

    (function(){
      var msg = 'fso is an object';

      var actual = typeof fso;
      var expected = 'object';

      assert.same(actual, expected, msg);
    }());
  });

  test('BVBA', function (assert) {

    (function(){
      var msg = 'BVBA is an object';

      var actual = typeof BVBA;
      var expected = 'object';

      assert.same(actual, expected, msg);
    }());
  });

  test('BVBA.getApp', function (assert) {

    (function(){
      var msg = 'BVBA.getApp(\'Word\') returns a Microsoft Word application object';

      var app = BVBA.getApp('Word');
      var actual = app.Name;
      app.Quit();

      var expected = 'Microsoft Word';

      assert.same(actual, expected, msg);
    }());

    (function(){
      var msg = 'BVBA.getApp(\'Excel\') returns a Microsoft Excel application object';

      var app = BVBA.getApp('Excel');
      var actual = app.Name;
      app.Quit();

      var expected = 'Microsoft Excel';

      assert.same(actual, expected, msg);
    }());

    (function(){
      var pathScript = WScript.ScriptFullName;
      var f = fso.GetFile(pathScript);

      var msg = 'BVBA.getParentFolderName is the parent folder name';

      var actual = BVBA.getParentFolderName();
      var expected = fso.GetParentFolderName(f);

      assert.same(actual, expected, msg);
    }());
  });
  
}());