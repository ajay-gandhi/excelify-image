'use strict';

const lwip     = require('lwip'),
      unzip    = require('unzip2'),
      archiver = require('archiver'),
      fs       = require('fs'),
      rimraf   = require('rimraf'),
      libxml   = require('libxmljs');

module.exports = (path) => {
  let suc = (c_map, image) => {
    let done = false;

    // Unzip template
    fs
      .createReadStream(__dirname + '/template.xlsx')
      .pipe(unzip.Extract({ path: 'output/' }))
      .on('close', () => {
        // Write styles
        fs.readFile(__dirname + '/output/xl/styles.xml', (err, data) => {
          if (err) throw err;

          let xml = libxml.parseXml(data.toString());
          let fills = xml.childNodes()[1];
          let xfs = xml.childNodes()[4];

          // For each unique color, add an XML element
          const colors   = Object.keys(c_map),
                n_colors = colors.length;
          for (let i = 0; i < n_colors; i++) {
            c_map[colors[i]] = i + 1;

            // Create fill
            let new_fill = fills.node('fill');

            let new_pfill = new_fill.node('patternFill');
            new_pfill.attr({ patternType: 'solid' });

            let new_fgc = new_fill.node('fgColor');
            new_fgc.attr({ rgb: colors[i] });

            let new_bgc = new_fill.node('bgColor');
            new_bgc.attr({ indexed: '64'});

            // Create xf
            let new_xf = xfs.node('xf').attr({
              numFmtId: 0,
              fontId: 0,
              fillId: i + 2,
              borderId: 0,
              xfId: 0,
              applyFill: 1
            });
          }

          fills.attr({ count: n_colors + 2 });
          xfs.attr({ count: n_colors + 1 });

          fs.writeFile(__dirname + '/output/xl/styles.xml', xml.toString(false), (err) => {
            if (err) throw err;

            console.log('done with styles');
            if (done) rezip_xlsx();
            else      done = true;
          });
        });

        // Write actual data
        fs.readFile(__dirname + '/output/xl/worksheets/sheet1.xml', (err, data) => {
          if (err) throw err;

          let xml = libxml.parseXml(data);

          // Fix dimensions
          let dimension = xml.childNodes()[0];
          dimension.attr({
            ref: 'A1:' + int_to_alpha(image[0].length - 1) + image[0].length
          });

          let sheet_format = xml.childNodes()[2];

          // Create custom width columns
          let cols = new libxml.Element(xml, 'cols');
          const ncols = image[0].length;
          cols.node('col').attr({
            min: 1,
            max: ncols,
            width: 2.5,
            customWidth: 1
          });

          sheet_format.addNextSibling(cols);

          // Create cells
          let sheet_data = xml.childNodes()[4];
          const nrows = image.length;
          for (let j = 0; j < nrows; j++) {
            // Create row
            let new_row = sheet_data.node('row');
            new_row.attr({
              r: j + 1,
              spans: '1:' + ncols
            });

            // Add cells
            for (let i = 0; i < ncols; i++) {
              let new_cell = new_row.node('c');
              new_cell.attr({
                r: int_to_alpha(i) + (j + 1),
                s: c_map[image[j][i]]
              });
            }
          }

          // Write file and finish up
          fs.writeFile(__dirname + '/output/xl/worksheets/sheet1.xml', xml.toString(false), (err) => {
            if (err) throw err;

            console.log('done with data');
            if (done) rezip_xlsx();
            else      done = true;
          });
        });
      });
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
  var output = fs.createWriteStream('./output.xlsx');
  var archive = archiver('zip');

  console.log('zipping');

  output.on('close', () => {
    console.log("Created output file");
    //rimraf(__dirname + '/output', () => {
    //  console.log('Output file saved to ./output.xlsx');
    //});
  });

  archive.pipe(output);
  archive.glob(__dirname + '/output/*');
  archive.finalize();
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
  lwip.open(path, function(err, image) {
    if (err) return failure_func('Error loading image: ' + err);

    // Setup options
    let options = {
      fit:     opts.fit     ? opts.fit               : 'original',
      width:   opts.width   ? parseInt(opts.width)   : image.width(),
      height:  opts.height  ? parseInt(opts.height)  : image.height()
    }

    let new_dims = calculate_dims(image, options);

    // Resize to requested dimensions
    image.resize(new_dims[0], new_dims[1], function (err, image) {
      if (err) return failure_func('Error resizing image: ' + err);

      // Keep list of all colors in the image
      let color_map = {};
      let c_list = [];

      // Get and convert pixels
      let i, j, pixel, color;
      const height = image.height(),
            width  = image.width();
      for (j = 0; j < height; j++) {

        // Add new array
        c_list.push([]);

        for (i = 0; i < width; i++) {
          pixel = image.getPixel(i, j);
          color = rgb_to_hex(pixel.r, pixel.g, pixel.b);

          // Transparency of pixel
          if (pixel.a) color = (Math.ceil(pixel.a * 2.55)).toString(16) + color;
          else         color = 'FF' + color;

          color = color.toUpperCase();
          if (!(color in color_map)) color_map[color] = color;
          c_list[j].push(color);
        }
      }

      success_func(color_map, c_list);

    });
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
  return '' + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

/**
 * Calculates the new dimensions of the image, given the options.
 *
 * @param [Image]  img  - The image (only width and height props needed)
 * @param [Object] opts - The options object
 *
 * @returns [Array] An array of the format [width, height]
 */
function calculate_dims(img, opts) {
  switch (opts.fit) {

    // Scale down by width
    case 'width':
      return [opts.width, img.height() * (opts.width / img.width())];

    // Scale down by height
    case 'height':
      return [img.width() * (opts.height / img.height()), opts.height];

    // Scale by width and height (ignore aspect ratio)
    case 'none':
      return [opts.width, opts.height];

    // Scale down to fit inside box matching width/height of options
    case 'box':
      var w_ratio = img.width()  / opts.width,
          h_ratio = img.height() / opts.height,
          neww, newh;

      if (w_ratio > h_ratio) {
          newh = Math.round(img.height() / w_ratio);
          neww = opts.width;
      } else {
          neww = Math.round(img.width() / h_ratio);
          newh = opts.height;
      }
      return [neww, newh];

    // Don't change width/height
    // Also the default in case of bad argument
    case 'original':
    default:
      // Let them know, but continue
      if (opts.fit !== 'original')
        console.error('Invalid option "fit", assuming "original"');

      return [img.width(), img.height()];

  }
}

