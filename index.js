const jimp     = require('jimp'),
      fs       = require('fs'),
      xl       = require("excel4node");

const PIXEL_WIDTH = 2.67;
const PIXEL_HEIGHT = 16;

module.exports = (path) => {

  let suc = (c_map, image) => {
    // Create a new instance of a Workbook class
    const wb = new xl.Workbook();

    // Add Worksheets to the workbook
    const ws = wb.addWorksheet('Sheet 1');

    // Create reusable styles for all colors
    Object.keys(c_map).forEach((color) => {
      c_map[color] = {
        fill: {
          type: "pattern",
          patternType: "solid",
          fgColor: color,
        },
      };
    });

    let i, j;
    for (i = 0; i < image.length; i++) {
      ws.column(i + 1).setWidth(PIXEL_WIDTH);
      for (j = 0; j < image[i].length; j++) {
        if (j === 0) {
          ws.row(j + 1).setHeight(PIXEL_HEIGHT);
        }
        ws.cell(i + 1, j + 1).style(c_map[image[i][j]]);
      }
    }

    wb.write('Excel.xlsx');
  }
  let fai = (msg) => {
    console.log(msg);
  }
  asciify_core(path, {}, suc, fai);
}

/**
 * Zips the files back into the xlsx file
 */
function rezip_xlsx() {
  console.log("Please manually zip for now:");
  console.log("  zip -r output.xlsx xl/ docProps/ ...");
  // var output = fs.createWriteStream('./output.xlsx');
  // var archive = archiver('zip');

  // console.log('zipping');

  // output.on('close', () => {
    // console.log("Created output file");
    // //rimraf(__dirname + '/output', () => {
    // //  console.log('Output file saved to ./output.xlsx');
    // //});
  // });

  // archive.pipe(output);
  // archive.glob(__dirname + '/output/*');
  // archive.finalize();
}

/**
 * The module's core functionality.
 *
 * @param  [string]   path      - The full path to the image to be asciified
 * @param  [Object]   opts      - The options object
 * @param  [Function] success_func - Callback if asciification succeeds
 * @param  [Function] failure_func - Callback if asciification fails
 *
 * @returns [void]
 */
function asciify_core(path, opts, success_func, failure_func) {
  // First open image to get initial properties
  jimp.read(path, function(err, image) {
    if (err) return failure_func('Error loading image: ' + err);

    // Keep list of all colors in the image
    let color_map = {};
    let c_list = [];

    // Get and convert pixels
    let i, j, pixel, color;
    for (j = 0; j < image.bitmap.height; j++) {

      // Add new array
      c_list.push([]);

      for (i = 0; i < image.bitmap.width; i++) {
        pixel = jimp.intToRGBA(image.getPixelColor(i, j));
        color = rgb_to_hex(pixel.r, pixel.g, pixel.b);

        // Transparency of pixel
        // color = 'FF' + color;
        // if (pixel.a) color = (Math.ceil(pixel.a * 2.55)).toString(16) + color;
        // else         color = 'FF' + color;

        color = color.toUpperCase();
        if (!(color in color_map)) color_map[color] = color;
        c_list[j].push(color);
      }
    }

    success_func(color_map, c_list);
  });
}

/**
 * Converts an integer to a string:
 *   1  -> A
 *   6  -> F
 *   28 -> AC
 *   55 -> BD
 */
const letters = 'abcdefghijklmnopqrstuvwxyz';
function int_to_alpha(val) {
  // Calculate number of digits
  val += 1;
  let digits = 0;
  let x = 1;
  while (val >= x) {
    digits++;
    val -= x;
    x *= 26;
  }

  // Now do base conversion
  let s = '';
  for (let i = 0; i < digits; i++) {
    s = letters.charAt(val % 26) + s;
    val = Math.floor(val / 26);
  }

  return s.toUpperCase();
}

/**
 * Converts RGB values to a hex string
 */
function rgb_to_hex (r, g, b) {
  return '#' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

