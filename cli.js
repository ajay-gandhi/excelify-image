
const excelify = require("./index");

const args = process.argv.slice(2);

const parseScale = (str) => {
  const parts = str.split("/");
  if (parts.length > 1) {
    return parseInt(parts[0], 10) / parseInt(parts[1], 10);
  } else {
    return Number(parts[0], 10);
  }
};

const { path, scale } = args.reduce((memo, arg) => {
  if (memo.nextIsScale) {
    return {
      ...memo,
      nextIsScale: false,
      scale: parseScale(arg),
    };
  } else if (arg.startsWith("--scale=")) {
    // --scale=x syntax
    return {
      ...memo,
      scale: parseScale(arg.slice(8)),
    };
  } else if (arg === "--scale" || arg === "-s") {
    // --scale x or -s x syntax
    return {
      ...memo,
      nextIsScale: true,
    };
  } else {
    // Assume it's the path
    return {
      ...memo,
      path: arg,
    };
  }
}, {
  path: false,
  nextIsScale: false,
});

if (!path) {
  console.log("Must enter a file to excelify");
  process.exit(1);
}
console.log(path, scale);
// excelify(path, { scale });

