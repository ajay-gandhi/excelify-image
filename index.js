const jimp = require("jimp");
const fs = require("fs");
const xl = require("excel4node");

const PIXEL_WIDTH = 2.67;
const PIXEL_HEIGHT = 16;

module.exports = async (path) => {
  const { colorList, pixelColorMap } = await parseImage(path);

  // Create a workbook with one sheet
  const workbook = new xl.Workbook();
  const worksheet = workbook.addWorksheet("Sheet 1");

  // Create one style for each color
  Object.keys(colorList).forEach((color) => {
    colorList[color] = {
      fill: {
        type: "pattern",
        patternType: "solid",
        fgColor: color,
      },
    };
  });

  // Write cells and adjust col width / row height
  let i, j;
  for (i = 0; i < pixelColorMap.length; i++) {
    worksheet.column(i + 1).setWidth(PIXEL_WIDTH);
    for (j = 0; j < pixelColorMap[i].length; j++) {
      if (j === 0) {
        worksheet.row(j + 1).setHeight(PIXEL_HEIGHT);
      }
      worksheet.cell(i + 1, j + 1).style(colorList[pixelColorMap[i][j]]);
    }
  }

  workbook.write("output.xlsx");
}

/**
 * Given a path to an image, reads the image and returns a hash of all the
 * colors used in the image, and a 2D array mapping pixels to their colors.
 *
 * 2D array is formatted like so:
 *   [
 *     [(0, 0), (0, 1), (0, 2)],
 *     [(1, 0), (1, 1), (1, 2)],
 *     ...
 *   ]
 */
const parseImage = async (path) => {
  return new Promise((resolve, reject) => {
    jimp.read(path, (err, image) => {
      if (err) return reject(err);

      // Maintain a hash of all colors used in the image
      // And the 2D array we'll return
      const colorList = {};
      const pixelColorMap = [];

      let i, j;
      for (j = 0; j < image.bitmap.height; j++) {
        pixelColorMap.push([]);
        for (i = 0; i < image.bitmap.width; i++) {
          const color = rgbToHex(jimp.intToRGBA(image.getPixelColor(i, j))).toUpperCase();
          if (!colorList[color]) colorList[color] = color;
          pixelColorMap[j].push(color);
        }
      }

      resolve({ colorList, pixelColorMap });
    });
  });
};

/**
 * Converts given RGB values to a hex string
 */
const rgbToHex = (p) => {
  return "#" + ((1 << 24) + (p.r << 16) + (p.g << 8) + p.b).toString(16).slice(1);
}

