// global object used to compute some values only once
// see https://stackoverflow.com/a/70074204

const G = {};

const addGetter_ = (name, value, obj=G) => {
  Object.defineProperty(obj, name, {
    enumerable: true,
    configurable: true,
    get() {
      // delete getter and replace with actual return value
      delete this[name];
      return (this[name] = value());
    },
  });
  return obj;
};

// global variables in G
[
  ['ss', () => SpreadsheetApp.getActive()],
  ['ui', () => SpreadsheetApp.getUi()],
  ['uiAvailable', () => uiAvailable()],
].forEach(([n, v]) => addGetter_(n, v));
